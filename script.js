/* ===================================================
   PANORAMA AGRO — script.js
   =================================================== */

// ── Fonte dos dados ────────────────────────────────
// PASSO 1: Coloque o ID da sua planilha do Google Sheets abaixo.
// Para obter: abra a planilha → copie a URL → pegue o trecho entre /d/ e /edit
// PASSO 2: A planilha DEVE ser pública: Compartilhar → "Qualquer pessoa com o link" → Visualizador
const SHEET_ID = '13LX4-3YPRaAXu9E40uXDDwx6N-BDUzQE';

// Proxies CORS usados como fallback quando a URL direta falha (ex: GitHub Pages)
function buildSheetUrls() {
  var base = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?format=xlsx';
  return [
    base,
    'https://corsproxy.io/?url=' + encodeURIComponent(base),
    'https://api.allorigins.win/raw?url=' + encodeURIComponent(base),
    'https://cors-anywhere.herokuapp.com/' + base
  ];
}

// Intervalo de atualização automática
var AUTO_REFRESH_INTERVAL = 60 * 60 * 1000; // 1 hora

// ── Estado global ──────────────────────────────────
var DATA = { panorama: [], saldo: [], fluxo: [], lavoura: [], contratos: [], tradeAlcance: [] };
var PARSED = { panorama: null, lavoura: [], contratos: [], fluxo: [], tradeAlcance: [] };
var FILTRO_EMPRESA = 'TODAS';
var FILTRO_SAFRA_GLOBAL = 'TODAS'; // novo filtro global de safra
var SAFRA_SOJA  = null;
var SAFRA_MILHO = null;
var SAFRA_CONTRATO = null;

var chartSaldo, chartExposicao, chartRecPag, chartCultura, chartSojaBar;
var fluxoDataTable, contratosDataTable;
var darkMode = false;

// ── Indicadores de mercado (cache) ────────────────
var MARKET_DATA = {
  ptax: null, ptaxSub: '—',
  cdi: null, cdiSub: '—',
  selic: null, selicSub: '—',
  sojaChicago: null, sojaChicagoSub: '—',
  milhoChicago: null, milhoChicagoSub: '—',
  clima: { pluvio: '—', umidade: '—', temp: '—', previsao: '—', sub: '—' }
};

// ── Boot ───────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function() {
  updateClock();
  setInterval(updateClock, 1000);
  initCharts();
  setupNavHighlight();
  autoLoadFromGoogleDrive();
  setInterval(autoLoadFromGoogleDrive, AUTO_REFRESH_INTERVAL);
  fetchMarketData();
  setInterval(fetchMarketData, 15 * 60 * 1000);
});

// ── Relógio ────────────────────────────────────────
function updateClock() {
  var now   = new Date();
  var dateEl = document.getElementById('headerDate');
  if (dateEl) {
    var opts = { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' };
    dateEl.textContent = now.toLocaleDateString('pt-BR', opts) +
      ' · ' + now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  }
}

// ── Carga automática do Google Sheets ──────────────
async function autoLoadFromGoogleDrive() {
  var badge    = document.getElementById('gfBadge');
  var statusEl = document.getElementById('driveStatus');
  var loadText = document.querySelector('.loading-text');

  if (badge)    badge.textContent = 'Atualizando...';
  if (statusEl) statusEl.textContent = '⏳ Conectando ao Google Sheets...';
  if (loadText) loadText.textContent = 'Conectando ao Google Sheets...';

  showLoading(true);

  var urls = buildSheetUrls();
  var lastError = null;

  for (var i = 0; i < urls.length; i++) {
    var url   = urls[i];
    var label = i === 0 ? 'direto' : 'proxy ' + i;

    try {
      if (statusEl) statusEl.textContent = '⏳ Tentando ' + label + '...';
      if (loadText) loadText.textContent = 'Buscando dados (' + label + ')...';

      // Fetch com timeout manual (compatível com todos os browsers)
      var buf = await fetchWithTimeout(url, 25000);

      if (!buf || buf.byteLength < 200) throw new Error('Resposta vazia');

      if (loadText) loadText.textContent = 'Lendo planilha...';

      var wb = XLSX.read(buf, { type: 'array', cellDates: true });

      processWorkbook(wb);

      var hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
      if (badge)    badge.textContent = '✅ Sheets · ' + hora;
      if (statusEl) statusEl.textContent = '✅ Atualizado às ' + hora + ' via ' + label;

      showLoading(false);
      return; // sucesso

    } catch (err) {
      console.warn('[autoLoad] Falha via ' + label + ':', err.message);
      lastError = err;
    }
  }

  // Todas as tentativas falharam
  console.error('[autoLoad] Todas as tentativas falharam:', lastError);
  if (badge)    badge.textContent = '⚠ Erro — verifique se a planilha é pública';
  if (statusEl) statusEl.textContent = '❌ Não foi possível conectar. Verifique se a planilha é pública ou importe manualmente.';
  showLoading(false);
}

// Fetch com timeout manual (sem AbortSignal.timeout para compatibilidade)
function fetchWithTimeout(url, ms) {
  return new Promise(function(resolve, reject) {
    var done = false;
    var timer = setTimeout(function() {
      if (!done) { done = true; reject(new Error('Timeout após ' + ms + 'ms')); }
    }, ms);

    fetch(url, { method: 'GET', cache: 'no-store' })
      .then(function(res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.arrayBuffer();
      })
      .then(function(buf) {
        if (!done) { done = true; clearTimeout(timer); resolve(buf); }
      })
      .catch(function(err) {
        if (!done) { done = true; clearTimeout(timer); reject(err); }
      });
  });
}

// Processa workbook (usado tanto no auto-load quanto no import manual)
function processWorkbook(wb) {
  DATA.panorama    = sheetToArray(wb, 'PANORAMA');
  DATA.lavoura     = sheetToArray(wb, 'LAVOURA');
  DATA.contratos   = sheetToArray(wb, 'CONTRATOS');
  DATA.fluxo       = sheetToArray(wb, 'FLUXO');
  DATA.tradeAlcance = sheetToArray(wb, 'TRADE_ALCANCE');

  PARSED.panorama    = parsePanorama();
  PARSED.lavoura     = parseLavoura();
  PARSED.contratos   = parseContratos();
  PARSED.fluxo       = parseFluxo();
  PARSED.tradeAlcance = parseTradeAlcance();

  FILTRO_EMPRESA = 'TODAS';
  buildFilterChips(PARSED.panorama.empresas);
  buildSafraGlobalChips();
  initSafraTabs();
  initContratoSafraTabs();
  refreshAll();
}

// ── Sidebar / UI controls ─────────────────────────
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('collapsed');
  document.getElementById('mainContent').classList.toggle('expanded');
}
function toggleDark() {
  darkMode = !darkMode;
  document.documentElement.setAttribute('data-dark', darkMode);
  if (PARSED.panorama) refreshAll();
}
function toggleFullscreen() {
  if (!document.fullscreenElement) document.documentElement.requestFullscreen();
  else document.exitFullscreen();
}
function exportPDF() { window.print(); }

// ── Nav highlight ─────────────────────────────────
function setupNavHighlight() {
  var obs = new IntersectionObserver(function(entries) {
    entries.forEach(function(e) {
      if (e.isIntersecting) {
        document.querySelectorAll('.nav-item').forEach(function(n) { n.classList.remove('active'); });
        var t = document.querySelector('.nav-item[data-section="' + e.target.id + '"]');
        if (t) t.classList.add('active');
      }
    });
  }, { threshold: 0.25 });
  document.querySelectorAll('.section').forEach(function(s) { obs.observe(s); });
}

// ── Formatadores ──────────────────────────────────
var fBRL = function(v) {
  return (v == null || isNaN(v)) ? '—' :
    new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(v);
};
var fBRL2 = function(v) {
  return (v == null || isNaN(v)) ? '—' :
    new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v);
};
var fUSD = function(v) {
  return (v == null || isNaN(v)) ? '—' :
    new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(v);
};
var fNum = function(v, d) {
  d = d || 0;
  return (v == null || isNaN(v)) ? '—' :
    new Intl.NumberFormat('pt-BR', { maximumFractionDigits: d }).format(v);
};
var fPct = function(v) {
  if (v == null || isNaN(v)) return '—%';
  return v.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '%';
};
function fDate(v) {
  if (!v) return '—';
  if (v instanceof Date) return v.toLocaleDateString('pt-BR');
  if (typeof v === 'string') { var d = new Date(v); if (!isNaN(d)) return d.toLocaleDateString('pt-BR'); }
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000)).toLocaleDateString('pt-BR');
  return String(v);
}
function parseNum(v) {
  if (v == null || v === '' || v === '-') return 0;
  if (typeof v === 'number') return v;
  return parseFloat(String(v).replace(/\./g, '').replace(',', '.')) || 0;
}

// Formata valores de ORIGEM: se numérico, exibe decimal PT-BR (sem símbolo de moeda); senão texto
function formatOrigem(v) {
  if (!v && v !== 0) return '—';
  var s = String(v).trim();
  if (s === '' || s === '-' || s === '—') return '—';
  // Tenta interpretar como número
  var raw = s.replace(/\s/g, '');
  var n;
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(raw)) {
    n = parseFloat(raw.replace(/\./g, '').replace(',', '.'));
  } else if (/^\d+(,\d*)?$/.test(raw)) {
    n = parseFloat(raw.replace(',', '.'));
  } else if (/^\d+(\.\d+)?$/.test(raw)) {
    n = parseFloat(raw);
  } else {
    n = NaN;
  }
  if (!isNaN(n)) {
    return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return s;
}

// ── Leitura Excel (importação manual) ─────────────
function loadExcel(input) {
  var file = input.files[0];
  if (!file) return;

  // Limpa o input para permitir reimportar o mesmo arquivo
  input.value = '';

  var loadText = document.querySelector('.loading-text');
  var statusEl = document.getElementById('driveStatus');
  if (loadText) loadText.textContent = 'Lendo arquivo...';
  if (statusEl) statusEl.textContent = '⏳ Importando arquivo local...';
  showLoading(true);

  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      if (!window.XLSX) throw new Error('Biblioteca XLSX não carregada. Aguarde e tente novamente.');
      var buf = e.target.result;
      var wb  = XLSX.read(buf, { type: 'array', cellDates: true });

      // Verifica se as abas necessárias existem
      var required = ['PANORAMA', 'LAVOURA', 'CONTRATOS'];
      var missing  = required.filter(function(s) { return !wb.Sheets[s]; });
      if (missing.length > 0) {
        throw new Error('Aba(s) não encontrada(s): ' + missing.join(', ') +
          '. Verifique se o arquivo é a planilha correta.');
      }

      processWorkbook(wb);

      if (statusEl) statusEl.textContent = '✅ Arquivo importado com sucesso';
      var badge = document.getElementById('gfBadge');
      if (badge) badge.textContent = '✅ Arquivo local';

    } catch (err) {
      console.error('[loadExcel]', err);
      var msg = 'Erro ao ler a planilha:\n\n' + err.message;
      if (statusEl) statusEl.textContent = '❌ ' + err.message;
      alert(msg);
    }
    showLoading(false);
  };

  reader.onerror = function() {
    alert('Erro ao ler o arquivo. Tente novamente.');
    showLoading(false);
  };

  reader.readAsArrayBuffer(file);
}

function sheetToArray(wb, name) {
  var sheet = wb.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
}

function showLoading(show) {
  var el = document.getElementById('loadingOverlay');
  if (el) el.style.display = show ? 'flex' : 'none';
}

// ── PARSERS ───────────────────────────────────────
function parsePanorama() {
  var rows = DATA.panorama;
  var empresas = [];
  var NOMES_VALIDOS = ['AGRODIMER','KUMMEL','RANCHO ALEGRE','COSTA BEBER'];
  for (var i = 6; i <= 9; i++) {
    var r = rows[i];
    if (!r || !r[0]) continue;
    var nome = String(r[0]).trim();
    if (!NOMES_VALIDOS.some(function(v) { return nome.toUpperCase() === v.toUpperCase(); })) continue;
    empresas.push({
      nome:      nome,
      dataBase:  r[1],
      saldo:     parseNum(r[2]),
      dataRomp:  r[3],
      diasRomp:  parseNum(r[4]),
      recUSD:    parseNum(r[5]),
      pagUSD:    parseNum(r[6]),
      exposicao: parseNum(r[7]),
    });
  }
  return { empresas: empresas };
}

function parseLavoura() {
  var rows = DATA.lavoura;
  var items = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r || !r[0]) continue;
    items.push({
      empresa:     String(r[0]).trim(),
      cultura:     String(r[1] || '').trim().toUpperCase(),
      safra:       String(r[2] || '').trim(),
      area:        parseNum(r[3]),
      rendimento:  parseNum(r[4]),
      producao:    parseNum(r[5]),
      vendas:      parseNum(r[6]),
      saldo:       parseNum(r[7]),
      vendidoUSD:  parseNum(r[8]),
      vendidoBRL:  parseNum(r[9]),
      precoMedUSD: parseNum(r[10]),
      precoMedBRL: parseNum(r[11]),
    });
  }
  return items;
}

function parseContratos() {
  var rows = DATA.contratos;
  var items = [];
  if (!rows || !rows.length) return items;

  // ── Normalização ─────────────────────────────────────────────────
  function norm(s) {
    return String(s || '').trim()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/\s+/g, ' ');
  }

  // ── Detecta linha de cabeçalho ────────────────────────────────────
  var headerRow = -1;
  for (var h = 0; h < Math.min(rows.length, 5); h++) {
    var r0 = rows[h];
    if (!r0) continue;
    var nonEmpty = r0.filter(function(c) { return c != null && String(c).trim() !== ''; });
    var textCount = nonEmpty.filter(function(c) { return isNaN(parseFloat(String(c))); }).length;
    if (textCount >= Math.ceil(nonEmpty.length * 0.5)) { headerRow = h; break; }
  }
  var startRow = headerRow >= 0 ? headerRow + 1 : 1;
  var header   = headerRow >= 0 ? rows[headerRow].map(norm) : [];

  // ── Busca coluna por alias exato ou parcial ────────────────────────
  function col() {
    var aliases = Array.prototype.slice.call(arguments).map(norm);
    // Busca exata
    for (var a = 0; a < aliases.length; a++) {
      var idx = header.indexOf(aliases[a]);
      if (idx >= 0) return idx;
    }
    // Busca parcial (header contém o alias)
    for (var b = 0; b < aliases.length; b++) {
      for (var c = 0; c < header.length; c++) {
        if (header[c] && header[c].indexOf(aliases[b]) >= 0) return c;
      }
    }
    return -1;
  }

  // ── Mapeamento exato baseado na planilha real ──────────────────────
  // Header real: Nome da Origem | Tipo Ent/Saida | Cliente/Fornecedor |
  //              Número Contrato | Qt. Contrato (SC) | Qt. Entregue (SC) |
  //              Qt. A Entregar (SC) | Qt. Faturada (SC) | Qt. A Faturar (SC) |
  //              Qt. Fixada (SC) | Cultura | Safra
  var iEmpresa  = col('NOME DA ORIGEM', 'NOME ORIGEM', 'ORIGEM', 'EMPRESA', 'FAZENDA', 'ENTIDADE', 'PRODUTOR');
  var iTipo     = col('TIPO ENT/SAIDA', 'TIPO ENT SAIDA', 'TIPO', 'TIPO CONTRATO', 'MODALIDADE');
  var iCliente  = col('CLIENTE/FORNECEDOR', 'CLIENTE FORNECEDOR', 'CLIENTE', 'COMPRADOR', 'TRADING', 'FORNECEDOR');
  var iNumCont  = col('NUMERO CONTRATO', 'NUMERO DO CONTRATO', 'NUM CONTRATO', 'CONTRATO', 'NRO CONTRATO', 'NR CONTRATO', 'COD CONTRATO');
  var iQtCont   = col('QT. CONTRATO (SC)', 'QT CONTRATO (SC)', 'QT CONTRATO', 'QTD CONTRATO', 'QUANTIDADE CONTRATO', 'SACAS CONTRATO', 'QT TOTAL');
  var iQtEntr   = col('QT. ENTREGUE (SC)', 'QT ENTREGUE (SC)', 'QT ENTREGUE', 'QTD ENTREGUE', 'ENTREGUE');
  var iQtAEntr  = col('QT. A ENTREGAR (SC)', 'QT A ENTREGAR (SC)', 'QT A ENTREGAR', 'QTD A ENTREGAR', 'A ENTREGAR', 'SALDO A ENTREGAR');
  var iQtFat    = col('QT. FATURADA (SC)', 'QT FATURADA (SC)', 'QT FATURADA', 'QTD FATURADA', 'FATURADO');
  var iQtAFat   = col('QT. A FATURAR (SC)', 'QT A FATURAR (SC)', 'QT A FATURAR', 'QTD A FATURAR', 'A FATURAR', 'SALDO A FATURAR');
  var iQtFixed  = col('QT. FIXADA (SC)', 'QT FIXADA (SC)', 'QT FIXADA', 'QTD FIXADA', 'FIXADO', 'FIXADA');
  var iCultura  = col('CULTURA', 'PRODUTO', 'GRAO', 'GRAOS', 'COMMODITY');
  var iSafra    = col('SAFRA', 'ANO SAFRA', 'PERIODO', 'ANO');
  // Colunas extras (ainda não presentes na planilha — prontas para quando forem adicionadas)
  var iPrecoMed = col('PRECO MEDIO', 'PRECO MED', 'PRECO/SC', 'R$/SC', 'PRECO SACA', 'PRECO UNITARIO', 'PRECO');
  var iMoeda    = col('MOEDA', 'CURRENCY', 'MOEDA CONTRATO', 'TIPO MOEDA');
  var iValorCont= col('VALOR CONTRATO', 'VL CONTRATO', 'VALOR TOTAL', 'VL TOTAL', 'TOTAL', 'VALOR');
  var iPrazoEmb = col('PRAZO EMBARQUE', 'DT EMBARQUE', 'DATA EMBARQUE', 'EMBARQUE', 'PRAZO', 'PERIODO EMBARQUE');
  var iDtPgto   = col('DATA PGTO', 'DT PGTO', 'DATA PAGAMENTO', 'DT PAGAMENTO', 'VENCIMENTO', 'DATA VENC');
  var iFrete    = col('FRETE', 'TIPO FRETE', 'MODALIDADE FRETE', 'CONDICAO FRETE');
  var iFazenda  = col('FAZENDA', 'PROPRIEDADE', 'LOCAL', 'LOCAL ENTREGA', 'UNIDADE');

  // Fallback por posição ordinal
  if (iEmpresa  < 0) iEmpresa  = 0;
  if (iTipo     < 0) iTipo     = 1;
  if (iCliente  < 0) iCliente  = 2;
  if (iNumCont  < 0) iNumCont  = 3;
  if (iQtCont   < 0) iQtCont   = 4;
  if (iQtEntr   < 0) iQtEntr   = 5;
  if (iQtAEntr  < 0) iQtAEntr  = 6;
  if (iQtFat    < 0) iQtFat    = 7;
  if (iQtAFat   < 0) iQtAFat   = 8;
  if (iQtFixed  < 0) iQtFixed  = 9;
  if (iCultura  < 0) iCultura  = 10;
  if (iSafra    < 0) iSafra    = 11;
  // Colunas extras: NÃO usa fallback de posição — só preenche se encontradas no header

  console.log('[CONTRATOS] Header detectado (linha ' + headerRow + '):', header.join(' | '));
  console.log('[CONTRATOS] EMP=' + iEmpresa + ' TIPO=' + iTipo + ' CLI=' + iCliente +
    ' CONT=' + iNumCont + ' QT=' + iQtCont + ' ENT=' + iQtEntr + ' AENT=' + iQtAEntr +
    ' FAT=' + iQtFat + ' AFAT=' + iQtAFat + ' FIX=' + iQtFixed +
    ' CULT=' + iCultura + ' SAFRA=' + iSafra);

  // ── Leitura das linhas ────────────────────────────────────────────
  for (var i = startRow; i < rows.length; i++) {
    var r = rows[i];
    if (!r) continue;
    var empVal = r[iEmpresa];
    if (empVal == null || String(empVal).trim() === '') continue;
    var empStr = norm(empVal);
    if (empStr === 'TOTAL' || empStr === 'TOTAIS' || empStr === 'SUBTOTAL') continue;

    var qtContrato  = parseNum(r[iQtCont]);
    var qtEntregue  = parseNum(r[iQtEntr]);
    var qtAEntregar = parseNum(r[iQtAEntr]);
    var qtFaturada  = parseNum(r[iQtFat]);
    var qtAFaturar  = parseNum(r[iQtAFat]);
    var qtFixada    = parseNum(r[iQtFixed]);
    var pctAtend    = qtContrato > 0 ? (qtEntregue / qtContrato) * 100 : 0;

    function safeStr(idx) {
      return (idx >= 0 && r[idx] != null) ? String(r[idx]).trim() : '';
    }
    function safeNum(idx) {
      return (idx >= 0 && r[idx] != null) ? parseNum(r[idx]) : null;
    }
    function safeDate(idx) {
      return (idx >= 0 && r[idx] != null) ? r[idx] : null;
    }

    items.push({
      empresa:       String(empVal).trim(),
      tipo:          safeStr(iTipo),
      cliente:       safeStr(iCliente),
      numContrato:   String(r[iNumCont] || '').trim(),
      qtContrato:    qtContrato,
      qtEntregue:    qtEntregue,
      qtAEntregar:   qtAEntregar,
      qtFaturada:    qtFaturada,
      qtAFaturar:    qtAFaturar,
      qtFixada:      qtFixada,
      cultura:       safeStr(iCultura).toUpperCase(),
      safra:         safeStr(iSafra),
      pctAtend:      pctAtend,
      // Colunas extras (preenchidas quando adicionadas na planilha)
      precoMedio:    safeNum(iPrecoMed),
      moeda:         safeStr(iMoeda),
      valorContrato: safeNum(iValorCont),
      prazoEmbarque: safeDate(iPrazoEmb),
      dataPgto:      safeDate(iDtPgto),
      frete:         safeStr(iFrete),
      fazenda:       safeStr(iFazenda),
    });
  }

  console.log('[CONTRATOS] Linhas carregadas:', items.length);
  return items;
}

function parseFluxo() {
  var rows = DATA.fluxo;
  if (!rows || !rows.length) return [];

  // Header normalizado: sem acentos, uppercase, sem espaços duplos
  function norm(s) {
    return String(s || '').trim()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/\s+/g, ' ');
  }

  var header = rows[0] ? rows[0].map(norm) : [];

  // Busca coluna por lista de aliases (retorna primeiro encontrado)
  function col() {
    var aliases = Array.prototype.slice.call(arguments);
    for (var a = 0; a < aliases.length; a++) {
      var n = norm(aliases[a]);
      var idx = header.indexOf(n);
      if (idx >= 0) return idx;
    }
    return -1;
  }

  var iData      = col('DATA', 'DT', 'DT LANCAMENTO', 'DATA LANCAMENTO', 'DATA MOVIMENTO');
  var iEmpresa   = col('EMPRESA', 'FAZENDA', 'ENTIDADE', 'CIA');
  var iNome      = col('NOME', 'DESCRICAO', 'HISTORICO', 'DESCR', 'NOME LANCAMENTO');
  var iEntOrig   = col('ENTRADA_ORIGEM', 'ENTRADA ORIGEM', 'ORIGEM ENTRADA', 'TIPO ENTRADA', 'CATEGORIA ENTRADA', 'ORIGEM_ENTRADA');
  var iSaiOrig   = col('SAIDA_ORIGEM', 'SAIDA ORIGEM', 'ORIGEM SAIDA', 'TIPO SAIDA', 'CATEGORIA SAIDA', 'ORIGEM_SAIDA');
  var iEntradas  = col('ENTRADAS', 'ENTRADA', 'CREDITO', 'CREDITOS', 'VL ENTRADA', 'VALOR ENTRADA', 'RECEITA');
  var iSaidas    = col('SAIDAS', 'SAIDA', 'DEBITO', 'DEBITOS', 'VL SAIDA', 'VALOR SAIDA', 'DESPESA');
  var iSaldoDia  = col('SALDO DIA', 'SALDO_DIA', 'SALDO DO DIA', 'SALDO DIARIO');
  var iSaldoBanc = col('SALDO BANCARIO', 'SALDO_BANCARIO', 'SALDO BANCO', 'SALDO FINAL', 'SALDO AC');

  // Debug no console para facilitar diagnóstico em campo
  console.log('[FLUXO] Header detectado:', header);
  console.log('[FLUXO] Colunas mapeadas: DATA=' + iData + ' EMP=' + iEmpresa + ' NOME=' + iNome +
    ' ENT_ORIG=' + iEntOrig + ' SAI_ORIG=' + iSaiOrig +
    ' ENTRADAS=' + iEntradas + ' SAIDAS=' + iSaidas +
    ' SALDO_DIA=' + iSaldoDia + ' SALDO_BANC=' + iSaldoBanc);

  var items = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r) continue;
    var data = iData >= 0 ? r[iData] : null;
    if (!data) continue;

    var parsedDate = null;
    if (data instanceof Date) {
      parsedDate = new Date(data.getTime());
    } else if (typeof data === 'number') {
      parsedDate = new Date(Math.round((data - 25569) * 86400 * 1000));
    } else if (typeof data === 'string') {
      var d2 = new Date(data);
      if (!isNaN(d2)) parsedDate = d2;
    }
    if (!parsedDate || isNaN(parsedDate.getTime())) continue;

    items.push({
      data:      parsedDate,
      empresa:   iEmpresa  >= 0 ? String(r[iEmpresa]  || '').trim() : '',
      nome:      iNome     >= 0 ? String(r[iNome]     || '').trim() : '',
      entOrig:   iEntOrig  >= 0 ? String(r[iEntOrig]  || '').trim() : '',
      saiOrig:   iSaiOrig  >= 0 ? String(r[iSaiOrig]  || '').trim() : '',
      entradas:  iEntradas >= 0 ? parseNum(r[iEntradas]) : 0,
      saidas:    iSaidas   >= 0 ? parseNum(r[iSaidas])   : 0,
      saldoDia:  iSaldoDia >= 0 ? parseNum(r[iSaldoDia]) : 0,
      saldoBanc: iSaldoBanc >= 0 ? parseNum(r[iSaldoBanc]) : 0,
    });
  }
  return items;
}

// ── PARSER: TRADE_ALCANCE ─────────────────────────
function parseTradeAlcance() {
  var rows = DATA.tradeAlcance;
  if (!rows || !rows.length) return {};

  function norm(s) {
    return String(s || '').trim()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/\s+/g, ' ');
  }
  var header = rows[0] ? rows[0].map(norm) : [];

  function col() {
    var aliases = Array.prototype.slice.call(arguments).map(norm);
    for (var a = 0; a < aliases.length; a++) {
      var idx = header.indexOf(aliases[a]);
      if (idx >= 0) return idx;
    }
    return -1;
  }

  var iContrato  = col('CONTRATO','NRO CONTRATO','Nº CONTRATO','NUM CONTRATO','NUMERO CONTRATO','NUM_CONTRATO','NUMERO_CONTRATO','N CONTRATO','NR CONTRATO');
  if (iContrato < 0) iContrato = 0;

  var iEmpresa   = col('EMPRESA','FAZENDA','ENTIDADE','CIA');
  var iCultura   = col('CULTURA','PRODUTO','GRAO');
  var iCliente   = col('CLIENTE','COMPRADOR','DESTINO','TRADINGS');
  var iPrecoSaca = col('PRECO SACA','PRECO/SACA','VALOR SACA','VL SACA','PRECO SC','R$/SC','PRECO_SACA','PRECO SC (R$)','PRECO','PRECO MEDIO','PRECO MED');
  var iTipoFrete = col('TIPO FRETE','FRETE','MODALIDADE FRETE','TIPO_FRETE','CONDICAO FRETE','COND FRETE');
  var iQtSacas   = col('QT SACAS','QUANTIDADE','QT CONTRATO','SACAS','VOLUME','QTD','SACAS CONTRATO');
  var iValorTotal= col('VALOR TOTAL','VL TOTAL','VALOR CONTRATO','TOTAL','VL TOTAL CONTRATO','VALOR_TOTAL');
  var iDtPgto    = col('DATA PAGAMENTO','DT PAGAMENTO','VENCIMENTO','DT VENC','DATA VENC','DATA_PAGAMENTO','PRAZO PAGAMENTO','DATA PGTO','DT PGTO');
  var iPrazoEmb  = col('PRAZO EMBARQUE','DT EMBARQUE','DATA EMBARQUE','EMBARQUE','PRAZO_EMBARQUE','DT ENTREGA','PRAZO ENTREGA');
  var iSafra     = col('SAFRA','ANO SAFRA','PERIODO');
  var iObs       = col('OBS','OBSERVACAO','OBSERVACOES','NOTA');
  var iMoeda     = col('MOEDA','CURRENCY','MOEDA CONTRATO');
  var iFazenda   = col('FAZENDA','PROPRIEDADE','LOCAL ENTREGA');

  console.log('[TRADE] Header:', header);
  console.log('[TRADE] Map: CONTRATO='+iContrato+' EMP='+iEmpresa+' CULTURA='+iCultura+' CLIENTE='+iCliente+
    ' PRECO='+iPrecoSaca+' FRETE='+iTipoFrete+' QT='+iQtSacas+
    ' VL_TOTAL='+iValorTotal+' DT_PGTO='+iDtPgto+' PRAZO_EMB='+iPrazoEmb);

  var map = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r) continue;
    var raw = r[iContrato];
    var numContrato = (raw != null && String(raw).trim() !== '')
      ? String(raw).trim().toUpperCase()
      : 'SEM NRO';

    // SEM NRO: índice único para não sobrescrever
    var key = numContrato === 'SEM NRO' ? 'SEM NRO__' + i : numContrato;

    map[key] = {
      numContrato: numContrato,
      empresa:    iEmpresa    >= 0 ? String(r[iEmpresa]    || '').trim() : '',
      cultura:    iCultura    >= 0 ? String(r[iCultura]    || '').trim().toUpperCase() : '',
      cliente:    iCliente    >= 0 ? String(r[iCliente]    || '').trim() : '',
      precoSaca:  iPrecoSaca  >= 0 ? parseNum(r[iPrecoSaca])  : null,
      tipoFrete:  iTipoFrete  >= 0 ? String(r[iTipoFrete]  || '').trim() : '',
      qtSacas:    iQtSacas    >= 0 ? parseNum(r[iQtSacas])    : null,
      valorTotal: iValorTotal >= 0 ? parseNum(r[iValorTotal]) : null,
      dtPgto:     iDtPgto     >= 0 ? r[iDtPgto]  : null,
      prazoEmb:   iPrazoEmb   >= 0 ? r[iPrazoEmb] : null,
      safra:      iSafra      >= 0 ? String(r[iSafra]     || '').trim() : '',
      obs:        iObs        >= 0 ? String(r[iObs]       || '').trim() : '',
      moeda:      iMoeda      >= 0 ? String(r[iMoeda]     || '').trim() : '',
      fazenda:    iFazenda    >= 0 ? String(r[iFazenda]   || '').trim() : '',
    };
  }
  return map;
}
function buildFilterChips(empresas) {
  var wrap = document.getElementById('gfChips');
  if (!wrap) return;
  wrap.innerHTML =
    '<button class="gf-chip active" data-emp="TODAS" onclick="setEmpresaFiltro(\'TODAS\',this)">' +
    '<i class="fas fa-layer-group"></i> Todas</button>';
  empresas.forEach(function(e) {
    var btn = document.createElement('button');
    btn.className = 'gf-chip';
    btn.dataset.emp = e.nome;
    btn.innerHTML = '<i class="fas fa-building"></i> ' + e.nome;
    btn.onclick = function() { setEmpresaFiltro(e.nome, btn); };
    wrap.appendChild(btn);
  });
}

function buildSafraGlobalChips() {
  var wrap = document.getElementById('gfSafraChips');
  var wrapOuter = document.getElementById('gfSafraWrap');
  if (!wrap || !wrapOuter) return;

  // Coleta todas as safras disponíveis (lavoura + contratos)
  var safras = [];
  (PARSED.lavoura || []).forEach(function(l) {
    if (l.safra && safras.indexOf(l.safra) < 0) safras.push(l.safra);
  });
  (PARSED.contratos || []).forEach(function(c) {
    if (c.safra && safras.indexOf(c.safra) < 0) safras.push(c.safra);
  });
  safras.sort();

  if (!safras.length) { wrapOuter.style.display = 'none'; return; }

  wrapOuter.style.display = 'flex';
  wrap.innerHTML = '<button class="gf-safra-chip active" data-safra="TODAS" onclick="setSafraGlobal(\'TODAS\',this)">Todas</button>';
  safras.forEach(function(s) {
    var label = s === '2026/2026' ? '2025/2026' : s; // normaliza label
    var btn = document.createElement('button');
    btn.className = 'gf-safra-chip';
    btn.dataset.safra = s;
    btn.textContent = label;
    btn.onclick = (function(safra, b) { return function() { setSafraGlobal(safra, b); }; })(s, btn);
    wrap.appendChild(btn);
  });

  FILTRO_SAFRA_GLOBAL = 'TODAS';
}

function setSafraGlobal(safra, btn) {
  FILTRO_SAFRA_GLOBAL = safra;
  document.querySelectorAll('.gf-safra-chip').forEach(function(c) { c.classList.remove('active'); });
  btn.classList.add('active');
  refreshAll();
}

function setEmpresaFiltro(emp, btn) {
  FILTRO_EMPRESA = emp;
  document.querySelectorAll('.gf-chip').forEach(function(c) { c.classList.remove('active'); });
  btn.classList.add('active');
  refreshAll();
}

function empresasFiltradas() {
  var all = PARSED.panorama ? PARSED.panorama.empresas : [];
  return FILTRO_EMPRESA === 'TODAS' ? all : all.filter(function(e) { return e.nome === FILTRO_EMPRESA; });
}

function lavouraFiltrada() {
  var all = PARSED.lavoura || [];
  if (FILTRO_EMPRESA !== 'TODAS') all = all.filter(function(l) { return l.empresa === FILTRO_EMPRESA; });
  if (FILTRO_SAFRA_GLOBAL !== 'TODAS') all = all.filter(function(l) { return l.safra === FILTRO_SAFRA_GLOBAL; });
  return all;
}

function contratosFiltrados() {
  var all = PARSED.contratos || [];
  if (FILTRO_EMPRESA !== 'TODAS') all = all.filter(function(c) { return c.empresa === FILTRO_EMPRESA; });

  // Filtro safra: prioridade → select local → global → aba tabs
  var selSafra   = document.getElementById('filtroContratoSafraExtra');
  var safraExtra = selSafra ? selSafra.value : '';

  if (safraExtra) {
    all = all.filter(function(c) { return c.safra === safraExtra; });
  } else if (FILTRO_SAFRA_GLOBAL !== 'TODAS') {
    all = all.filter(function(c) { return c.safra === FILTRO_SAFRA_GLOBAL; });
  } else if (SAFRA_CONTRATO) {
    // só filtra se SAFRA_CONTRATO não for vazio
    all = all.filter(function(c) { return c.safra === SAFRA_CONTRATO; });
  }
  // se SAFRA_CONTRATO === '' → sem filtro de safra, mostra tudo

  var selCultura    = document.getElementById('filtroContratoCultura');
  var culturaFiltro = selCultura ? selCultura.value : '';
  if (culturaFiltro) all = all.filter(function(c) { return c.cultura === culturaFiltro; });

  return all;
}

// ── SAFRA TABS ─────────────────────────────────────
function initSafraTabs() {
  var lav       = PARSED.lavoura;
  var sojaSafras  = [];
  var milhoSafras = [];
  var allSafras   = [];

  lav.forEach(function(l) {
    if (l.cultura === 'SOJA'  && sojaSafras.indexOf(l.safra) < 0)  sojaSafras.push(l.safra);
    if (l.cultura === 'MILHO' && milhoSafras.indexOf(l.safra) < 0) milhoSafras.push(l.safra);
    if (allSafras.indexOf(l.safra) < 0) allSafras.push(l.safra);
  });
  sojaSafras.sort();  milhoSafras.sort();  allSafras.sort();

  SAFRA_SOJA  = sojaSafras[0]  || null;
  SAFRA_MILHO = milhoSafras[0] || null;

  buildSafraTabs('sojaSafraTabs',  sojaSafras,  'SOJA');
  buildSafraTabs('milhoSafraTabs', milhoSafras, 'MILHO');

  ['filterSafraCultura', 'filterSafraEmpresa'].forEach(function(id) {
    var sel = document.getElementById(id);
    if (!sel) return;
    var cur = sel.value;
    sel.innerHTML = '<option value="">Todas as Safras</option>';
    allSafras.forEach(function(s) {
      var opt = document.createElement('option');
      opt.value = s; opt.textContent = s;
      if (s === cur) opt.selected = true;
      sel.appendChild(opt);
    });
  });

  var allContratoSafras = [];
  (PARSED.contratos || []).forEach(function(c) {
    if (allContratoSafras.indexOf(c.safra) < 0) allContratoSafras.push(c.safra);
  });
  allContratoSafras.sort();

  var selExt = document.getElementById('filtroContratoSafraExtra');
  if (selExt) {
    var cur2 = selExt.value;
    selExt.innerHTML = '<option value="">Todas</option>';
    allContratoSafras.forEach(function(s) {
      var opt = document.createElement('option');
      opt.value = s; opt.textContent = s;
      if (s === cur2) opt.selected = true;
      selExt.appendChild(opt);
    });
  }
}

function buildSafraTabs(containerId, safras, cultura) {
  var wrap = document.getElementById(containerId);
  if (!wrap) return;
  wrap.innerHTML = '';
  safras.forEach(function(s, i) {
    var btn = document.createElement('button');
    btn.className = 'safra-tab' + (i === 0 ? ' active' : '');
    // Para MILHO: 2026/2026 exibe como 2025/2026
    btn.textContent = (cultura === 'MILHO' && s === '2026/2026') ? '2025/2026' : s;
    btn.onclick = function() {
      wrap.querySelectorAll('.safra-tab').forEach(function(b) { b.classList.remove('active'); });
      btn.classList.add('active');
      if (cultura === 'SOJA')  { SAFRA_SOJA  = s; renderGrainTable('sojaBody',  lavouraFiltrada().filter(function(l) { return l.cultura === 'SOJA'  && l.safra === s; })); }
      if (cultura === 'MILHO') { SAFRA_MILHO = s; renderGrainTable('milhoBody', lavouraFiltrada().filter(function(l) { return l.cultura === 'MILHO' && l.safra === s; })); }
    };
    wrap.appendChild(btn);
  });
}

// ── CONTRATOS SAFRA TABS ──────────────────────────
function initContratoSafraTabs() {
  var allSafras = [];
  (PARSED.contratos || []).forEach(function(c) {
    if (c.safra && allSafras.indexOf(c.safra) < 0) allSafras.push(c.safra);
  });
  allSafras.sort();
  // Por padrão, sem filtro de safra — mostra todos os contratos
  SAFRA_CONTRATO = '';
  buildContratoSafraTabs(allSafras);

  // Garante botão Agrupar no estado correto
  var btn = document.getElementById('btnAgrupar');
  if (btn) {
    btn.classList.toggle('active', AGRUPAR_EMPRESAS);
    btn.innerHTML = AGRUPAR_EMPRESAS
      ? '<i class="fas fa-list"></i> Desagrupar'
      : '<i class="fas fa-layer-group"></i> Agrupar por Empresa';
  }
}

function buildContratoSafraTabs(safras) {
  var wrap = document.getElementById('contratoSafraTabs');
  if (!wrap) return;
  wrap.innerHTML = '';

  // Botão "Todas" — default ativo
  var btnTodas = document.createElement('button');
  btnTodas.className = 'safra-tab active';
  btnTodas.textContent = 'Todas';
  btnTodas.onclick = function() {
    wrap.querySelectorAll('.safra-tab').forEach(function(b) { b.classList.remove('active'); });
    btnTodas.classList.add('active');
    SAFRA_CONTRATO = '';
    renderContratosTable();
  };
  wrap.appendChild(btnTodas);

  safras.forEach(function(s) {
    var btn = document.createElement('button');
    btn.className = 'safra-tab';
    // Milho: 2026/2026 exibe como 2025/2026
    btn.textContent = s === '2026/2026' ? '2025/2026' : s;
    btn.onclick = function() {
      wrap.querySelectorAll('.safra-tab').forEach(function(b) { b.classList.remove('active'); });
      btn.classList.add('active');
      SAFRA_CONTRATO = s;
      renderContratosTable();
    };
    wrap.appendChild(btn);
  });
}

// ── REFRESH COMPLETO ──────────────────────────────
function refreshAll() {
  if (!PARSED.panorama) return;
  var emps = empresasFiltradas();
  var lav  = lavouraFiltrada();

  renderKPIs(emps, lav);
  renderFluxoTable(emps);
  renderProximosRecebimentos();
  renderFluxo30Dias();
  renderAlertas(emps);
  renderAllCharts(emps, lav);
  renderGrains(lav);
  renderContratosTable();
  renderIndicadores(emps, lav);

  var now   = new Date();
  var badge = FILTRO_EMPRESA === 'TODAS' ? 'Todas as empresas' : FILTRO_EMPRESA;
  var el    = document.getElementById('lastUpdated');
  if (el) el.textContent = 'Atualizado ' +
    now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }) + ' · ' + badge;
}

// ── KPI CARDS ─────────────────────────────────────
function renderKPIs(emps, lav) {
  var saldoTotal = emps.reduce(function(s,e) { return s+e.saldo; }, 0);
  var totalRec   = emps.reduce(function(s,e) { return s+e.recUSD; }, 0);
  var totalPag   = emps.reduce(function(s,e) { return s+e.pagUSD; }, 0);
  var totalExp   = emps.reduce(function(s,e) { return s+e.exposicao; }, 0);
  var fluxo      = totalRec - totalPag;

  var safraAtivaSoja  = SAFRA_SOJA  || '2025/2026';
  var safraAtivaMilho = SAFRA_MILHO || '2025/2026';

  var soja  = lav.filter(function(l) { return l.cultura === 'SOJA'  && l.safra === safraAtivaSoja; });
  var milho = lav.filter(function(l) { return l.cultura === 'MILHO' && l.safra === safraAtivaMilho; });

  var totalVendasSoja  = soja.reduce(function(s,l) { return s+l.vendas; }, 0);
  var totalProdSoja    = soja.reduce(function(s,l) { return s+l.producao; }, 0);
  var totalDispSoja    = soja.reduce(function(s,l) { return s+l.saldo; }, 0);
  var totalVendasMilho = milho.reduce(function(s,l) { return s+l.vendas; }, 0);
  var pct = totalProdSoja > 0 ? (totalVendasSoja / totalProdSoja) * 100 : 0;

  setText('kpiSaldoTotal', fBRL(saldoTotal));
  setTrend('kpiSaldoTrend', 2.45, 'vs dia anterior');
  setText('kpiFluxo', 'USD ' + fUSD(fluxo));
  setTrend('kpiFluxoTrend', fluxo >= 0 ? 5.1 : -8.1, 'vs dia anterior');
  setText('kpiRecUSD', fUSD(totalRec));
  setTrend('kpiRecTrend', 5.31, 'vs dia anterior');
  setText('kpiPagUSD', fUSD(totalPag));
  setTrend('kpiPagTrend', 3.87, 'vs dia anterior');
  setText('kpiExposicao', 'USD ' + fUSD(totalExp));
  setTrend('kpiExpTrend', totalExp >= 0 ? 1.2 : -3.5, 'vs dia anterior');
  setText('kpiSafra', fNum(totalVendasSoja) + ' sc');
  setTrend('kpiSafraTrend', 3.10, 'vs dia anterior');
  setText('kpiSafraMilho', fNum(totalVendasMilho) + ' sc');
  setTrend('kpiSafraMilhoTrend', 2.80, 'vs dia anterior');
  setText('kpiSaldoDisp', fNum(totalDispSoja) + ' sc');
  setTrend('kpiSaldoDispTrend', -1.4, 'vs dia anterior');
  setText('kpiComercPct', fPct(pct));
  setTrend('kpiComercTrend', 2.15, 'vs dia anterior');
}

function setText(id, val) {
  var el = document.getElementById(id);
  if (el) el.textContent = val;
}
function setTrend(id, pct, label) {
  var el = document.getElementById(id);
  if (!el) return;
  var arrow = pct > 0 ? '▲' : pct < 0 ? '▼' : '→';
  var cls   = pct > 0 ? 'trend-up' : pct < 0 ? 'trend-down' : 'trend-neu';
  el.className = 'kpi-trend ' + cls;
  el.innerHTML = arrow + ' ' + Math.abs(pct).toLocaleString('pt-BR', {minimumFractionDigits:2, maximumFractionDigits:2}) + '% <span style="font-weight:400;opacity:.7">' + label + '</span>';
}

// ── TABELA FLUXO ──────────────────────────────────
function renderFluxoTable(emps) {
  var tbody = document.getElementById('fluxoBody');
  if (!tbody) return;

  // Destroi DataTable se existir
  if (fluxoDataTable) {
    try { fluxoDataTable.destroy(); } catch(e) {}
    fluxoDataTable = null;
  }
  tbody.innerHTML = '';

  emps.forEach(function(e) {
    var diasCls = e.diasRomp < 30 ? 'dias-vermelho' : e.diasRomp < 90 ? 'dias-amarelo' : 'dias-verde';
    var expCls  = e.exposicao < 0 ? 'exp-neg' : 'exp-pos';
    var tr = document.createElement('tr');
    tr.innerHTML =
      '<td><strong>' + e.nome + '</strong></td>' +
      '<td class="mono">' + fDate(e.dataBase) + '</td>' +
      '<td class="mono">' + fBRL(e.saldo) + '</td>' +
      '<td class="mono">' + fDate(e.dataRomp) + '</td>' +
      '<td><span class="' + diasCls + '">' + e.diasRomp + '</span></td>' +
      '<td class="mono">' + fUSD(e.recUSD) + '</td>' +
      '<td class="mono">' + fUSD(e.pagUSD) + '</td>' +
      '<td class="' + expCls + '">' + fUSD(e.exposicao) + '</td>';
    tbody.appendChild(tr);
  });

  if (typeof $ !== 'undefined' && $.fn.DataTable) {
    try {
      fluxoDataTable = $('#fluxoTable').DataTable({
        pageLength: 10,
        language: { search: 'Buscar:', lengthMenu: 'Exibir _MENU_', info: '_START_–_END_ de _TOTAL_',
                    paginate: { previous: '‹', next: '›' }, emptyTable: 'Sem dados' },
        order: [[4, 'asc']], responsive: true, destroy: true
      });
    } catch(e) { console.warn('DataTable init error:', e); }
  }
}

// ── PRÓXIMOS RECEBIMENTOS ─────────────────────────
function renderProximosRecebimentos() {
  var tbody  = document.getElementById('proxRecebBody');
  var banner = document.getElementById('proxRecebAlerta');
  var bannerMsg = document.getElementById('proxRecebAlertaMsg');
  var badge  = document.getElementById('proxRecebBadge');
  if (!tbody) return;

  var today  = new Date(); today.setHours(0,0,0,0);
  var limit  = new Date(today); limit.setDate(today.getDate() + 30);

  var fluxo = PARSED.fluxo || [];
  var emps  = FILTRO_EMPRESA === 'TODAS' ? null : FILTRO_EMPRESA;

  var items = fluxo.filter(function(r) {
    if (r.entradas <= 0) return false;
    if (emps && r.empresa !== emps) return false;
    var d = new Date(r.data); d.setHours(0,0,0,0);
    return d >= today && d <= limit;
  });

  items.sort(function(a, b) { return a.data - b.data; });

  var top10 = items.slice(0, 10);
  var totalValor = items.reduce(function(s, r) { return s + r.entradas; }, 0);
  var urgentes   = items.filter(function(r) {
    var d = new Date(r.data); d.setHours(0,0,0,0);
    var diff = (d - today) / 86400000;
    return diff <= 7;
  });

  // Alerta
  if (banner && bannerMsg) {
    if (urgentes.length > 0) {
      banner.style.display = 'flex';
      bannerMsg.textContent = urgentes.length + ' recebimento(s) esperado(s) nos próximos 7 dias · Total: ' + fBRL(urgentes.reduce(function(s,r){return s+r.entradas;},0));
      banner.className = 'alerta-receb-banner banner-warning';
    } else if (items.length === 0) {
      banner.style.display = 'flex';
      bannerMsg.textContent = 'Nenhum recebimento previsto nos próximos 30 dias.';
      banner.className = 'alerta-receb-banner banner-info';
    } else {
      banner.style.display = 'none';
    }
  }

  if (badge) badge.textContent = items.length + ' recebimento(s) · ' + fBRL(totalValor);

  tbody.innerHTML = '';
  if (!top10.length) {
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--text-muted)">Nenhum recebimento nos próximos 30 dias</td></tr>';
    return;
  }

  top10.forEach(function(r) {
    var d = new Date(r.data); d.setHours(0,0,0,0);
    var diasAte = Math.round((d - today) / 86400000);
    var urgCls  = diasAte <= 3 ? 'receb-urgente' : diasAte <= 7 ? 'receb-proximo' : '';
    var tr = document.createElement('tr');
    if (urgCls) tr.className = urgCls;
    tr.innerHTML =
      '<td class="mono">' + r.data.toLocaleDateString('pt-BR') +
        (diasAte === 0 ? ' <span class="receb-hoje">HOJE</span>' : diasAte === 1 ? ' <span class="receb-amanha">AMANHÃ</span>' : ' <span class="receb-dias">' + diasAte + 'd</span>') +
      '</td>' +
      '<td><strong>' + (r.empresa || '—') + '</strong></td>' +
      '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + r.nome + '">' + (r.nome || '—') + '</td>' +
      '<td><span class="origem-badge">' + formatOrigem(r.entOrig) + '</span></td>' +
      '<td class="mono" style="text-align:right;color:var(--green-mid);font-weight:700">' + fBRL(r.entradas) + '</td>';
    tbody.appendChild(tr);
  });

  if (items.length > 10) {
    var more = document.createElement('tr');
    more.innerHTML = '<td colspan="5" style="text-align:center;color:var(--text-muted);font-style:italic;font-size:0.8rem">+ ' + (items.length - 10) + ' registro(s) adicionais não exibidos</td>';
    tbody.appendChild(more);
  }
}

// ── FLUXO 30 DIAS ─────────────────────────────────
function renderFluxo30Dias() {
  var container = document.getElementById('fluxo30Container');
  if (!container) return;

  var today = new Date(); today.setHours(0,0,0,0);
  var limit = new Date(today); limit.setDate(today.getDate() + 30);

  var fluxo = PARSED.fluxo || [];
  var emps  = FILTRO_EMPRESA === 'TODAS' ? null : FILTRO_EMPRESA;

  var items = fluxo.filter(function(r) {
    if (emps && r.empresa !== emps) return false;
    var d = new Date(r.data); d.setHours(0,0,0,0);
    return d >= today && d <= limit;
  });
  items.sort(function(a, b) { return a.data - b.data; });

  if (!items.length) {
    container.innerHTML = '<div style="text-align:center;color:var(--text-muted);padding:2rem">Nenhum lançamento nos próximos 30 dias</div>';
    return;
  }

  // Agrupar por empresa
  var byEmp = {};
  items.forEach(function(r) {
    var emp = r.empresa || 'SEM EMPRESA';
    if (!byEmp[emp]) byEmp[emp] = [];
    byEmp[emp].push(r);
  });

  var html = '';
  Object.keys(byEmp).forEach(function(emp) {
    var rows = byEmp[emp];
    var totEnt = rows.reduce(function(s,r) { return s + r.entradas; }, 0);
    var totSai = rows.reduce(function(s,r) { return s + r.saidas;   }, 0);
    var totSal = rows.reduce(function(s,r) { return s + r.saldoDia; }, 0);
    var empKey = 'fluxo30_' + emp.replace(/\s/g,'_');

    html += '<div class="fluxo30-group">' +
      '<div class="fluxo30-group-header" onclick="toggleFluxo30Group(\'' + empKey + '\', this)">' +
        '<div class="fluxo30-emp-info">' +
          '<span class="group-toggle-icon"><i class="fas fa-chevron-right"></i></span>' +
          '<strong>' + emp + '</strong>' +
          '<span class="group-count">' + rows.length + ' lançamento(s)</span>' +
        '</div>' +
        '<div class="fluxo30-totais">' +
          '<span class="ft-ent"><i class="fas fa-arrow-down"></i> ' + fBRL(totEnt) + '</span>' +
          '<span class="ft-sai"><i class="fas fa-arrow-up"></i> ' + fBRL(totSai) + '</span>' +
          '<span class="ft-sal ' + (totSal >= 0 ? 'ft-sal-pos' : 'ft-sal-neg') + '"><i class="fas fa-scale-balanced"></i> ' + fBRL(totSal) + '</span>' +
        '</div>' +
      '</div>' +
      '<div class="fluxo30-body" id="' + empKey + '" style="display:none">' +
        '<div class="table-responsive">' +
          '<table class="table" style="width:100%;font-size:0.81rem">' +
            '<thead><tr>' +
              '<th>Data</th><th>Nome</th><th>Origem Entrada</th><th>Origem Saída</th>' +
              '<th style="text-align:right">Entradas</th><th style="text-align:right">Saídas</th>' +
              '<th style="text-align:right">Saldo Dia</th><th style="text-align:right">Saldo Bancário</th>' +
            '</tr></thead>' +
            '<tbody>' +
              rows.map(function(r) {
                var sCls = r.saldoDia >= 0 ? 'color:var(--green-mid)' : 'color:var(--red)';
                return '<tr>' +
                  '<td class="mono">' + r.data.toLocaleDateString('pt-BR') + '</td>' +
                  '<td>' + (r.nome || '—') + '</td>' +
                  '<td><span class="origem-badge">' + formatOrigem(r.entOrig) + '</span></td>' +
                  '<td><span class="origem-badge origem-sai">' + formatOrigem(r.saiOrig) + '</span></td>' +
                  '<td class="mono" style="text-align:right;color:var(--green-mid)">' + (r.entradas > 0 ? fBRL(r.entradas) : '—') + '</td>' +
                  '<td class="mono" style="text-align:right;color:var(--red)">'       + (r.saidas   > 0 ? fBRL(r.saidas)   : '—') + '</td>' +
                  '<td class="mono" style="text-align:right;' + sCls + '">' + fBRL(r.saldoDia) + '</td>' +
                  '<td class="mono" style="text-align:right">' + fBRL(r.saldoBanc) + '</td>' +
                '</tr>';
              }).join('') +
            '</tbody>' +
          '</table>' +
        '</div>' +
      '</div>' +
    '</div>';
  });

  container.innerHTML = html;
}

function toggleFluxo30Group(id, headerEl) {
  var body = document.getElementById(id);
  if (!body) return;
  var isOpen = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : '';
  var icon = headerEl.querySelector('.group-toggle-icon i');
  if (icon) icon.className = isOpen ? 'fas fa-chevron-right' : 'fas fa-chevron-down';
}

// ── GRAINS ────────────────────────────────────────
function renderGrains(lav) {
  var sojaItems  = lav.filter(function(l) { return l.cultura === 'SOJA'  && l.safra === (SAFRA_SOJA  || '2025/2026'); });
  var milhoItems = lav.filter(function(l) { return l.cultura === 'MILHO' && l.safra === (SAFRA_MILHO || '2025/2026'); });
  renderGrainTable('sojaBody',  sojaItems);
  renderGrainTable('milhoBody', milhoItems);
}

function renderGrainTable(tbodyId, items) {
  var tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  tbody.innerHTML = '';
  items.sort(function(a, b) { return a.empresa.localeCompare(b.empresa); });
  items.forEach(function(item) {
    var pct    = item.producao > 0 ? (item.vendas / item.producao) * 100 : 0;
    var pctCls = pct >= 85 ? 'fill-red' : pct >= 60 ? 'fill-yellow' : 'fill-green';
    var color  = pct >= 85 ? 'var(--red)' : 'var(--green-mid)';
    var tr = document.createElement('tr');
    tr.innerHTML =
      '<td><strong>' + item.empresa + '</strong></td>' +
      '<td class="mono">' + fNum(item.producao) + '</td>' +
      '<td class="mono">' + fNum(item.vendas) + '</td>' +
      '<td class="mono">' + (pct >= 85 ? '⚠ ' : '') + fNum(item.saldo) + '</td>' +
      '<td style="font-weight:700;color:' + color + '">' + fPct(pct) + '</td>' +
      '<td style="min-width:70px"><div class="progress-bar-wrap"><div class="progress-bar-fill ' + pctCls + '" style="width:' + Math.min(pct,100).toFixed(1) + '%"></div></div></td>';
    tbody.appendChild(tr);
  });
}

// ── CONTRATOS TABLE ───────────────────────────────
var AGRUPAR_EMPRESAS = true;

function toggleAgruparEmpresas() {
  AGRUPAR_EMPRESAS = !AGRUPAR_EMPRESAS;
  var btn = document.getElementById('btnAgrupar');
  if (btn) {
    btn.classList.toggle('active', AGRUPAR_EMPRESAS);
    btn.innerHTML = AGRUPAR_EMPRESAS
      ? '<i class="fas fa-list"></i> Desagrupar'
      : '<i class="fas fa-layer-group"></i> Agrupar por Empresa';
  }
  renderContratosTable();
}

function exportContratosExcel() {
  var items = contratosFiltrados();
  if (!items.length) { alert('Nenhum contrato para exportar.'); return; }

  var headers = [
    'Empresa','Cultura','Safra','Nº Contrato','Cliente / Fornecedor',
    'Preço Médio (R$/sc)','Moeda','Valor Contrato (R$)',
    'Prazo Embarque','Data Pagto','Frete','Fazenda',
    'Qt. Contrato (sc)','Qt. Entregue (sc)','Qt. A Entregar (sc)','% Atendimento'
  ];

  var rows = items.map(function(c) {
    return [
      c.empresa        || '',
      c.cultura        || '',
      c.safra          || '',
      c.numContrato    || '',
      c.cliente        || '',
      c.precoMedio     != null ? c.precoMedio     : '',
      c.moeda          || '',
      c.valorContrato  != null ? c.valorContrato  : '',
      c.prazoEmbarque  ? fDate(c.prazoEmbarque) : '',
      c.dataPgto       ? fDate(c.dataPgto)      : '',
      c.frete          || '',
      c.fazenda        || '',
      c.qtContrato     != null ? c.qtContrato     : '',
      c.qtEntregue     != null ? c.qtEntregue     : '',
      c.qtAEntregar    != null ? c.qtAEntregar    : '',
      c.pctAtend       != null ? (c.pctAtend / 100) : '',
    ];
  });

  var wb2 = XLSX.utils.book_new();
  var ws  = XLSX.utils.aoa_to_sheet([headers].concat(rows));
  ws['!cols'] = headers.map(function(h, i) { return { wch: i >= 12 ? 16 : 22 }; });
  XLSX.utils.book_append_sheet(wb2, ws, 'Contratos');
  var safra = (document.getElementById('filtroContratoSafraExtra') || {}).value || SAFRA_CONTRATO || 'contratos';
  XLSX.writeFile(wb2, 'contratos_' + safra.replace('/','_') + '.xlsx');
}

// ── JOIN: CONTRATOS (ERP) + TRADE_ALCANCE (Gerencial) ───────────────
// Retorna lista unificada: base = TRADE_ALCANCE, enriquecida com ERP onde numContrato bater
function buildContratosMerged() {
  var erpMap = {};
  (PARSED.contratos || []).forEach(function(c) {
    var key = c.numContrato; // já normalizado uppercase em parseContratos
    if (!erpMap[key]) erpMap[key] = [];
    erpMap[key].push(c);
  });

  var ta = PARSED.tradeAlcance || {};
  var merged = [];

  // Itera sobre TRADE_ALCANCE como fonte primária
  Object.keys(ta).forEach(function(key) {
    var td = ta[key];
    var nro = td.numContrato; // 'SEM NRO' ou número real

    // Aplica filtro de empresa se ativo
    var empFiltro = FILTRO_EMPRESA !== 'TODAS' ? FILTRO_EMPRESA : null;
    if (empFiltro && td.empresa && td.empresa.toUpperCase() !== empFiltro.toUpperCase()) {
      // tenta também pelo ERP
      var erpRows = erpMap[nro] || [];
      if (!erpRows.some(function(e) { return e.empresa === empFiltro; })) return;
    }

    // Busca linha(s) ERP correspondente
    var erpRows = (nro !== 'SEM NRO' ? erpMap[nro] : null) || [];

    if (erpRows.length > 0) {
      // Merge: uma linha por contrato ERP, com dados Trade completando
      erpRows.forEach(function(erp) {
        merged.push({
          // Identificação
          numContrato: nro,
          empresa:     erp.empresa  || td.empresa  || '—',
          cultura:     erp.cultura  || td.cultura  || '—',
          cliente:     erp.cliente  || td.cliente  || '—',
          safra:       erp.safra    || td.safra    || '—',
          // Dados ERP (quantitativos)
          qtContrato:  erp.qtContrato,
          qtEntregue:  erp.qtEntregue,
          qtAEntregar: erp.qtAEntregar,
          qtFaturada:  erp.qtFaturada,
          qtAFaturar:  erp.qtAFaturar,
          pctAtend:    erp.qtContrato > 0 ? (erp.qtEntregue / erp.qtContrato) * 100 : 0,
          hasErp: true,
          // Dados Gerenciais (Trade) — prioriza Trade, fallback ERP
          tipo:         td.tipoFrete  || erp.tipo         || '',
          precoMedio:   td.precoSaca  != null ? td.precoSaca  : erp.precoMedio,
          moeda:        td.moeda      || erp.moeda         || '',
          valorContrato: td.valorTotal != null ? td.valorTotal : erp.valorContrato,
          prazoEmbarque: td.prazoEmb  != null ? td.prazoEmb  : erp.prazoEmbarque,
          dataPgto:     td.dtPgto     != null ? td.dtPgto    : erp.dataPgto,
          frete:        td.tipoFrete  || erp.frete          || '',
          fazenda:      td.fazenda    || erp.fazenda         || '',
          obs:          td.obs        || '',
          // Campos legados (compat)
          precoSaca:   td.precoSaca,
          tipoFrete:   td.tipoFrete,
          qtSacasTrade:td.qtSacas,
          valorTotal:  td.valorTotal,
          dtPgto:      td.dtPgto,
          prazoEmb:    td.prazoEmb,
        });
      });
    } else {
      // Contrato só no gerencial, sem par ERP
      merged.push({
        numContrato: nro,
        empresa:     td.empresa  || '—',
        cultura:     td.cultura  || '—',
        cliente:     td.cliente  || '—',
        safra:       td.safra    || '—',
        qtContrato:  td.qtSacas  || null,
        qtEntregue:  null,
        qtAEntregar: null,
        qtFaturada:  null,
        qtAFaturar:  null,
        pctAtend:    null,
        hasErp: false,
        tipo:         td.tipoFrete  || '',
        precoMedio:   td.precoSaca,
        moeda:        td.moeda      || '',
        valorContrato: td.valorTotal,
        prazoEmbarque: td.prazoEmb,
        dataPgto:     td.dtPgto,
        frete:        td.tipoFrete  || '',
        fazenda:      td.fazenda    || '',
        obs:          td.obs        || '',
        // Campos legados
        precoSaca:   td.precoSaca,
        tipoFrete:   td.tipoFrete,
        qtSacasTrade:td.qtSacas,
        valorTotal:  td.valorTotal,
        dtPgto:      td.dtPgto,
        prazoEmb:    td.prazoEmb,
      });
    }
  });

  // Adiciona contratos ERP que não têm par no Trade (número não cadastrado gerencialmente)
  Object.keys(erpMap).forEach(function(nro) {
    var hasTrade = Object.keys(ta).some(function(k) { return ta[k].numContrato === nro; });
    if (hasTrade) return; // já incluído acima
    erpMap[nro].forEach(function(erp) {
      var empFiltro = FILTRO_EMPRESA !== 'TODAS' ? FILTRO_EMPRESA : null;
      if (empFiltro && erp.empresa !== empFiltro) return;
      merged.push({
        numContrato: nro,
        empresa:     erp.empresa  || '—',
        cultura:     erp.cultura  || '—',
        cliente:     erp.cliente  || '—',
        safra:       erp.safra    || '—',
        qtContrato:  erp.qtContrato,
        qtEntregue:  erp.qtEntregue,
        qtAEntregar: erp.qtAEntregar,
        qtFaturada:  erp.qtFaturada,
        qtAFaturar:  erp.qtAFaturar,
        pctAtend:    erp.qtContrato > 0 ? (erp.qtEntregue / erp.qtContrato) * 100 : 0,
        hasErp: true,
        tipo:         erp.tipo          || '',
        precoMedio:   erp.precoMedio    != null ? erp.precoMedio   : null,
        moeda:        erp.moeda         || '',
        valorContrato: erp.valorContrato != null ? erp.valorContrato : null,
        prazoEmbarque: erp.prazoEmbarque,
        dataPgto:     erp.dataPgto,
        frete:        erp.frete         || '',
        fazenda:      erp.fazenda       || '',
        obs:          '',
        precoSaca: null, tipoFrete: '', qtSacasTrade: null,
        valorTotal: null, dtPgto: null, prazoEmb: null,
      });
    });
  });

  // Aplica filtros de safra e cultura
  var selSafraExtra = document.getElementById('filtroContratoSafraExtra');
  var safraExtra = selSafraExtra ? selSafraExtra.value : '';
  if (safraExtra) merged = merged.filter(function(m) { return m.safra === safraExtra; });
  else if (SAFRA_CONTRATO) merged = merged.filter(function(m) { return !m.safra || m.safra === '' || m.safra === SAFRA_CONTRATO; });

  var selCultura = document.getElementById('filtroContratoCultura');
  var culturaFiltro = selCultura ? selCultura.value : '';
  if (culturaFiltro) merged = merged.filter(function(m) { return m.cultura === culturaFiltro; });

  return merged;
}

function renderContratosTable() {
  var tbody = document.getElementById('contratosBody');
  if (!tbody) return;
  tbody.innerHTML = '';
  var NCOL = 19;
  var dash = '<span style="color:var(--text-muted)">—</span>';

  var items = contratosFiltrados();

  if (!items.length) {
    tbody.innerHTML = '<tr><td colspan="' + NCOL + '" style="text-align:center;padding:2rem;color:var(--text-muted)">' +
      '<i class="fas fa-inbox" style="font-size:1.5rem;display:block;margin-bottom:.5rem"></i>' +
      'Nenhum contrato encontrado</td></tr>';
    return;
  }

  // ── Célula helpers ────────────────────────────────────────────────
  function cellNum(v)  { return v != null ? fNum(v, 0)  : dash; }
  function cellBRL(v)  { return v != null && v !== 0 ? fBRL(v) : dash; }
  function cellBRL2(v) { return v != null && v !== 0 ? fBRL2(v) : dash; }
  function cellDate(v) { return v ? fDate(v) : dash; }
  function cellStr(v)  { return v ? String(v) : dash; }

  // ── Sub-linha de contrato (compartilhado entre modos) ─────────────
  function buildSubRow(c, indent) {
    var pct    = c.pctAtend != null ? c.pctAtend : 0;
    var pctCls = pct >= 80 ? 'fill-green' : pct >= 40 ? 'fill-yellow' : 'fill-red';
    var pctClr = pct >= 80 ? 'var(--green-light)' : pct >= 40 ? 'var(--yellow)' : 'var(--red)';
    var cultCls = c.cultura === 'SOJA' ? 'cultura-soja' : c.cultura === 'MILHO' ? 'cultura-milho' : 'cultura-outro';
    var moedaBadge = c.moeda
      ? '<span style="font-size:.68rem;font-weight:700;padding:2px 7px;border-radius:20px;background:rgba(46,134,222,.13);color:#1d4ed8">' + c.moeda + '</span>'
      : dash;

    var tr = document.createElement('tr');
    tr.className = 'contrato-sub-row';
    tr.innerHTML =
      '<td style="' + (indent ? 'padding-left:2rem;' : '') + 'font-size:.77rem;white-space:nowrap">' + (c.empresa || '—') + '</td>' +
      '<td><span class="cultura-badge ' + cultCls + '">' + (c.cultura || '—') + '</span></td>' +
      '<td style="font-size:.74rem;color:var(--text-muted);white-space:nowrap">' + (c.safra || '—') + '</td>' +
      '<td class="mono" style="font-size:.75rem;white-space:nowrap">' + (c.numContrato || '—') + '</td>' +
      '<td style="min-width:180px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:.76rem" title="' + (c.cliente || '') + '">' + (c.cliente || '—') + '</td>' +
      '<td class="mono" style="text-align:right;white-space:nowrap">' + cellBRL2(c.precoMedio) + '</td>' +
      '<td style="text-align:center">' + moedaBadge + '</td>' +
      '<td class="mono" style="text-align:right;white-space:nowrap">' + cellBRL(c.valorContrato) + '</td>' +
      '<td class="mono" style="text-align:center;white-space:nowrap">' + cellDate(c.prazoEmbarque) + '</td>' +
      '<td class="mono" style="text-align:center;white-space:nowrap">' + cellDate(c.dataPgto) + '</td>' +
      '<td style="white-space:nowrap">' + (c.frete ? '<span class="frete-badge">' + c.frete + '</span>' : dash) + '</td>' +
      '<td style="font-size:.76rem;white-space:nowrap">' + cellStr(c.fazenda) + '</td>' +
      '<td class="mono" style="text-align:right">' + cellNum(c.qtContrato)  + '</td>' +
      '<td class="mono" style="text-align:right">' + cellNum(c.qtEntregue)  + '</td>' +
      '<td class="mono" style="text-align:right">' + cellNum(c.qtAEntregar) + '</td>' +
      '<td>' + (c.pctAtend != null
        ? '<div style="display:flex;align-items:center;gap:5px;min-width:85px">' +
          '<div class="progress-bar-wrap" style="flex:1"><div class="progress-bar-fill ' + pctCls +
          '" style="width:' + Math.min(pct, 100).toFixed(1) + '%"></div></div>' +
          '<span style="font-size:.73rem;font-weight:700;color:' + pctClr + '">' + fPct(pct) + '</span></div>'
        : dash) + '</td>';
    return tr;
  }

  // ── AGRUPADO por empresa ──────────────────────────────────────────
  if (AGRUPAR_EMPRESAS) {
    var byEmp = {};
    var empOrder = [];
    items.forEach(function(c) {
      var emp = c.empresa || '—';
      if (!byEmp[emp]) { byEmp[emp] = []; empOrder.push(emp); }
      byEmp[emp].push(c);
    });
    // preserva ordem original (já ordenada pelo sort de empresas no parser)
    empOrder.sort();

    empOrder.forEach(function(emp) {
      var contratos = byEmp[emp];
      var totC   = contratos.reduce(function(s,c){ return s+(c.qtContrato  ||0); },0);
      var totE   = contratos.reduce(function(s,c){ return s+(c.qtEntregue  ||0); },0);
      var totAE  = contratos.reduce(function(s,c){ return s+(c.qtAEntregar ||0); },0);
      var totF   = contratos.reduce(function(s,c){ return s+(c.qtFaturada  ||0); },0);
      var totAF  = contratos.reduce(function(s,c){ return s+(c.qtAFaturar  ||0); },0);
      var totFix = contratos.reduce(function(s,c){ return s+(c.qtFixada    ||0); },0);
      var totPct = totC > 0 ? (totE / totC * 100) : 0;
      var totPctCls = totPct >= 80 ? 'fill-green' : totPct >= 40 ? 'fill-yellow' : 'fill-red';

      // Linha de grupo
      var groupRow = document.createElement('tr');
      groupRow.className = 'contrato-group-row';
      groupRow.style.cursor = 'pointer';
      groupRow.innerHTML =
        '<td colspan="5" style="white-space:nowrap">' +
          '<span class="group-toggle-icon"><i class="fas fa-chevron-right"></i></span>' +
          '<strong>' + emp + '</strong>' +
          ' <span class="group-count">' + contratos.length + ' contrato(s)</span>' +
        '</td>' +
        '<td colspan="7" style="color:var(--text-muted);font-size:.72rem;font-style:italic">Clique para expandir</td>' +
        '<td class="mono" style="text-align:right"><strong>' + fNum(totC,0)   + '</strong></td>' +
        '<td class="mono" style="text-align:right"><strong>' + fNum(totE,0)   + '</strong></td>' +
        '<td class="mono" style="text-align:right"><strong>' + fNum(totAE,0)  + '</strong></td>' +
        '<td><div style="display:flex;align-items:center;gap:6px;min-width:85px">' +
          '<div class="progress-bar-wrap" style="flex:1"><div class="progress-bar-fill ' + totPctCls +
          '" style="width:' + Math.min(totPct,100).toFixed(1) + '%"></div></div>' +
          '<span style="font-size:.74rem;font-weight:700">' + fPct(totPct) + '</span></div></td>';
      tbody.appendChild(groupRow);

      // Sub-linhas — começam escondidas (collapsed)
      var subRows = contratos.map(function(c) {
        var tr = buildSubRow(c, true);
        tr.style.display = 'none';
        tbody.appendChild(tr);
        return tr;
      });

      // Toggle
      groupRow.addEventListener('click', function() {
        var hidden = subRows[0] && subRows[0].style.display === 'none';
        subRows.forEach(function(r) { r.style.display = hidden ? '' : 'none'; });
        var icon = groupRow.querySelector('.group-toggle-icon i');
        if (icon) icon.className = hidden ? 'fas fa-chevron-down' : 'fas fa-chevron-right';
      });
    });

  } else {
    // ── DESAGRUPADO ───────────────────────────────────────────────
    items.forEach(function(c) {
      tbody.appendChild(buildSubRow(c, false));
    });
  }

  // ── Linha de TOTAIS ───────────────────────────────────────────────
  var tC  = items.reduce(function(s,c){ return s+(c.qtContrato  ||0); },0);
  var tE  = items.reduce(function(s,c){ return s+(c.qtEntregue  ||0); },0);
  var tAE = items.reduce(function(s,c){ return s+(c.qtAEntregar ||0); },0);
  var tF  = items.reduce(function(s,c){ return s+(c.qtFaturada  ||0); },0);
  var tAF = items.reduce(function(s,c){ return s+(c.qtAFaturar  ||0); },0);
  var tFix= items.reduce(function(s,c){ return s+(c.qtFixada    ||0); },0);
  var tVT = items.reduce(function(s,c){ return s+(c.valorContrato||0); },0);
  var tPct = tC > 0 ? (tE / tC * 100) : 0;
  var tPctCls = tPct >= 80 ? 'fill-green' : tPct >= 40 ? 'fill-yellow' : 'fill-red';

  var tfr = document.createElement('tr');
  tfr.className = 'contrato-total-row';
  tfr.innerHTML =
    '<td colspan="5"><strong>TOTAL — ' + items.length + ' contrato(s)</strong></td>' +
    '<td colspan="2" class="mono" style="text-align:right">' + (tVT > 0 ? '<strong>' + fBRL(tVT) + '</strong>' : '') + '</td>' +
    '<td colspan="5"></td>' +
    '<td class="mono" style="text-align:right"><strong>' + fNum(tC,0)   + '</strong></td>' +
    '<td class="mono" style="text-align:right"><strong>' + fNum(tE,0)   + '</strong></td>' +
    '<td class="mono" style="text-align:right"><strong>' + fNum(tAE,0)  + '</strong></td>' +
    '<td><div style="display:flex;align-items:center;gap:5px;min-width:85px">' +
      '<div class="progress-bar-wrap" style="flex:1"><div class="progress-bar-fill ' + tPctCls +
      '" style="width:' + Math.min(tPct,100).toFixed(1) + '%"></div></div>' +
      '<span style="font-size:.73rem;font-weight:700">' + fPct(tPct) + '</span></div></td>';
  tbody.appendChild(tfr);
}

function chartColors() {
  return darkMode
    ? { grid: 'rgba(255,255,255,.08)', text: '#8aad96', tick: '#8aad96' }
    : { grid: 'rgba(0,0,0,.06)',       text: '#6b7c74', tick: '#6b7c74' };
}

var PALETTE = ['#40916c','#2e86de','#e76f51','#7b2d8b','#0a9396','#f4a261','#e63946','#74c69d'];

function baseOpts() {
  var cc = chartColors();
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 500 },
    plugins: { legend: { labels: { color: cc.text, font: { family: 'Outfit', size: 11 }, boxWidth: 12 } } }
  };
}
function scaleOpts(yCallback) {
  var cc = chartColors();
  return {
    x: { ticks: { color: cc.tick, font: { size: 10 } }, grid: { color: cc.grid } },
    y: { ticks: { color: cc.tick, font: { size: 10 }, callback: yCallback }, grid: { color: cc.grid } }
  };
}

function initCharts() {
  chartSaldo = new Chart(document.getElementById('chartSaldo'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Saldo R$', data: [], backgroundColor: PALETTE, borderRadius: 5 }] },
    options: Object.assign({}, baseOpts(), { plugins: { legend: { display: false } },
      scales: scaleOpts(function(v) { return 'R$' + (Math.abs(v) >= 1e6 ? (v/1e6).toFixed(1)+'M' : (v/1e3).toFixed(0)+'k'); }) })
  });

  chartExposicao = new Chart(document.getElementById('chartExposicao'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Exposição USD', data: [], backgroundColor: [], borderRadius: 5 }] },
    options: Object.assign({}, baseOpts(), { plugins: { legend: { display: false } },
      scales: scaleOpts(function(v) { return (v/1e3).toFixed(0)+'k'; }) })
  });

  chartRecPag = new Chart(document.getElementById('chartRecPag'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Recebimentos', data: [], backgroundColor: 'rgba(64,145,108,.78)',  borderRadius: 4 },
      { label: 'Pagamentos',   data: [], backgroundColor: 'rgba(230,57,70,.75)',   borderRadius: 4 }
    ]},
    options: Object.assign({}, baseOpts(), { scales: scaleOpts(function(v) { return (v/1e6).toFixed(1)+'M'; }) })
  });

  chartCultura = new Chart(document.getElementById('chartCultura'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Comprometido', data: [], backgroundColor: 'rgba(230,57,70,.78)',  borderRadius: 4 },
      { label: 'Disponível',   data: [], backgroundColor: 'rgba(64,145,108,.82)', borderRadius: 4 }
    ]},
    options: Object.assign({}, baseOpts(), {
      scales: Object.assign({}, scaleOpts(function(v) { return (v/1e3).toFixed(0)+'k sc'; }),
        { x: { stacked: false, ticks: { color: chartColors().tick, font: { size: 10 } }, grid: { color: chartColors().grid } } }) })
  });

  chartSojaBar = new Chart(document.getElementById('chartSojaBar'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Produção Est.', data: [], backgroundColor: 'rgba(45,106,79,.6)',    borderRadius: 4 },
      { label: 'Vendas',        data: [], backgroundColor: 'rgba(64,145,108,.85)',  borderRadius: 4 },
      { label: 'Saldo Disp.',   data: [], backgroundColor: 'rgba(116,198,157,.75)', borderRadius: 4 }
    ]},
    options: Object.assign({}, baseOpts(), { scales: scaleOpts(function(v) { return (v/1e3).toFixed(0)+'k'; }) })
  });
}

function renderAllCharts(emps, lav) {
  var labels = emps.map(function(e) { return e.nome; });

  chartSaldo.data.labels = labels;
  chartSaldo.data.datasets[0].data = emps.map(function(e) { return e.saldo; });
  chartSaldo.data.datasets[0].backgroundColor = emps.map(function() { return 'rgba(64,145,108,.85)'; });
  chartSaldo.update();

  chartExposicao.data.labels = labels;
  chartExposicao.data.datasets[0].data = emps.map(function(e) { return e.exposicao; });
  chartExposicao.data.datasets[0].backgroundColor = emps.map(function(e) {
    return e.exposicao < 0 ? 'rgba(230,57,70,.8)' : 'rgba(46,134,222,.8)';
  });
  chartExposicao.update();

  chartRecPag.data.labels = labels;
  chartRecPag.data.datasets[0].data = emps.map(function(e) { return e.recUSD; });
  chartRecPag.data.datasets[1].data = emps.map(function(e) { return e.pagUSD; });
  chartRecPag.update();

  var culturas    = ['SOJA', 'MILHO'];
  var selCultura  = document.getElementById('filterSafraCultura');
  var safraCultura = selCultura ? selCultura.value : '';
  var lavCultura  = safraCultura ? lav.filter(function(l) { return l.safra === safraCultura; }) : lav;
  var comprom = culturas.map(function(c) { return lavCultura.filter(function(l) { return l.cultura===c; }).reduce(function(s,l) { return s+l.vendas; }, 0); });
  var dispon  = culturas.map(function(c) { return lavCultura.filter(function(l) { return l.cultura===c; }).reduce(function(s,l) { return s+l.saldo;  }, 0); });
  chartCultura.data.labels = culturas;
  chartCultura.data.datasets[0].data = comprom;
  chartCultura.data.datasets[1].data = dispon;
  chartCultura.update();

  var selEmpresa  = document.getElementById('filterSafraEmpresa');
  var safraEmpresa = selEmpresa ? selEmpresa.value : '';
  var safraAtiva  = safraEmpresa || SAFRA_SOJA || '2025/2026';
  var sojaItems   = safraEmpresa
    ? lav.filter(function(l) { return l.safra === safraEmpresa; })
    : lav.filter(function(l) { return l.cultura === 'SOJA' && l.safra === safraAtiva; });
  var empNames = [];
  sojaItems.forEach(function(l) { if (empNames.indexOf(l.empresa)<0) empNames.push(l.empresa); });
  chartSojaBar.data.labels = empNames;
  chartSojaBar.data.datasets[0].data = empNames.map(function(n) { return sojaItems.filter(function(l) { return l.empresa===n; }).reduce(function(s,l) { return s+l.producao; }, 0); });
  chartSojaBar.data.datasets[1].data = empNames.map(function(n) { return sojaItems.filter(function(l) { return l.empresa===n; }).reduce(function(s,l) { return s+l.vendas;   }, 0); });
  chartSojaBar.data.datasets[2].data = empNames.map(function(n) { return sojaItems.filter(function(l) { return l.empresa===n; }).reduce(function(s,l) { return s+l.saldo;    }, 0); });
  chartSojaBar.update();
}

// ── BUSCA DE DADOS DE MERCADO REAL ────────────────
async function fetchMarketData() {
  await Promise.allSettled([
    fetchPTAX(),
    fetchCDISELIC(),
    fetchChicagoPrices(),
    fetchClimatempo()
  ]);
  if (PARSED.panorama) renderIndicadores(empresasFiltradas(), lavouraFiltrada());
}

async function fetchPTAX() {
  try {
    var hoje = new Date();
    var fmt = function(d) {
      return String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0') + '-' + d.getFullYear();
    };
    for (var i = 1; i <= 5; i++) {
      var d = new Date(hoje);
      d.setDate(hoje.getDate() - i);
      var url = 'https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao=\'' + fmt(d) + '\'&$top=1&$format=json&$select=cotacaoCompra,cotacaoVenda';
      var res = await fetch(url);
      if (!res.ok) continue;
      var json = await res.json();
      var val  = json.value && json.value[0];
      if (val) {
        var ptax = ((val.cotacaoCompra + val.cotacaoVenda) / 2).toFixed(4);
        MARKET_DATA.ptax    = 'R$ ' + String(ptax).replace('.', ',');
        MARKET_DATA.ptaxSub = 'Compra R$' + String(val.cotacaoCompra.toFixed(4)).replace('.',',') + ' · Venda R$' + String(val.cotacaoVenda.toFixed(4)).replace('.',',');
        return;
      }
    }
    MARKET_DATA.ptax = '—'; MARKET_DATA.ptaxSub = 'Dado não disponível';
  } catch(e) {
    MARKET_DATA.ptax = '—'; MARKET_DATA.ptaxSub = 'Erro ao buscar PTAX';
  }
}

async function fetchCDISELIC() {
  try {
    var results = await Promise.allSettled([
      fetch('https://api.bcb.gov.br/dados/serie/bcdata.sgs.12/dados/ultimos/1?formato=json'),
      fetch('https://api.bcb.gov.br/dados/serie/bcdata.sgs.11/dados/ultimos/1?formato=json')
    ]);

    if (results[0].status === 'fulfilled' && results[0].value.ok) {
      var cdiData = await results[0].value.json();
      if (cdiData && cdiData[0]) {
        var cdiDiario = parseFloat(String(cdiData[0].valor).replace(',', '.'));
        var cdiAnual  = ((Math.pow(1 + cdiDiario / 100, 252) - 1) * 100).toFixed(2);
        MARKET_DATA.cdi    = cdiAnual.replace('.', ',') + '% a.a.';
        MARKET_DATA.cdiSub = 'Diário: ' + cdiData[0].valor + '% · ' + cdiData[0].data;
      }
    }

    if (results[1].status === 'fulfilled' && results[1].value.ok) {
      var selicData = await results[1].value.json();
      if (selicData && selicData[0]) {
        var selicDiario = parseFloat(String(selicData[0].valor).replace(',', '.'));
        var selicAnual  = ((Math.pow(1 + selicDiario / 100, 252) - 1) * 100).toFixed(2);
        MARKET_DATA.selic    = selicAnual.replace('.', ',') + '% a.a.';
        MARKET_DATA.selicSub = 'Diário: ' + selicData[0].valor + '% · ' + selicData[0].data;
      }
    }
  } catch(e) {
    MARKET_DATA.cdi = '—'; MARKET_DATA.cdiSub = 'Erro ao buscar CDI';
    MARKET_DATA.selic = '—'; MARKET_DATA.selicSub = 'Erro ao buscar SELIC';
  }
}

async function fetchChicagoPrices() {
  try {
    var results = await Promise.allSettled([
      fetch('https://query1.finance.yahoo.com/v8/finance/chart/ZS=F?interval=1d&range=5d'),
      fetch('https://query1.finance.yahoo.com/v8/finance/chart/ZC=F?interval=1d&range=5d')
    ]);

    if (results[0].status === 'fulfilled' && results[0].value.ok) {
      var j = await results[0].value.json();
      var c = j && j.chart && j.chart.result && j.chart.result[0];
      if (c) {
        var prices = (c.indicators && c.indicators.quote && c.indicators.quote[0] && c.indicators.quote[0].close) || [];
        var last   = prices.filter(function(v) { return v != null; }).pop();
        if (last) { MARKET_DATA.sojaChicago = last.toFixed(2) + ' USc/bu'; MARKET_DATA.sojaChicagoSub = 'CBOT · Yahoo Finance'; }
      }
    }

    if (results[1].status === 'fulfilled' && results[1].value.ok) {
      var j2 = await results[1].value.json();
      var c2 = j2 && j2.chart && j2.chart.result && j2.chart.result[0];
      if (c2) {
        var prices2 = (c2.indicators && c2.indicators.quote && c2.indicators.quote[0] && c2.indicators.quote[0].close) || [];
        var last2   = prices2.filter(function(v) { return v != null; }).pop();
        if (last2) { MARKET_DATA.milhoChicago = last2.toFixed(2) + ' USc/bu'; MARKET_DATA.milhoChicagoSub = 'CBOT · Yahoo Finance'; }
      }
    }
  } catch(e) {}

  if (!MARKET_DATA.sojaChicago)  { MARKET_DATA.sojaChicago  = '—'; MARKET_DATA.sojaChicagoSub  = 'Indisponível'; }
  if (!MARKET_DATA.milhoChicago) { MARKET_DATA.milhoChicago = '—'; MARKET_DATA.milhoChicagoSub = 'Indisponível'; }
}

async function fetchClimatempo() {
  try {
    var lat = -13.8278, lon = -54.1992;
    var url = 'https://api.open-meteo.com/v1/forecast?latitude=' + lat + '&longitude=' + lon +
      '&hourly=soil_moisture_0_to_1cm&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,weathercode&timezone=America%2FCuiaba&past_days=14&forecast_days=7';
    var res  = await fetch(url);
    if (!res.ok) throw new Error('Open-Meteo error');
    var data  = await res.json();
    var daily = data.daily;
    var today = new Date().toISOString().split('T')[0];

    var pluvioArr = daily.precipitation_sum || [];
    var datesArr  = daily.time || [];
    var pluvio14  = 0;
    datesArr.forEach(function(d, i) { if (d <= today && pluvioArr[i] != null) pluvio14 += pluvioArr[i]; });

    var tempMax = null, tempMin = null;
    datesArr.forEach(function(d, i) {
      if (d <= today) { tempMax = daily.temperature_2m_max[i]; tempMin = daily.temperature_2m_min[i]; }
    });
    var tempMedia = (tempMax != null && tempMin != null) ? ((tempMax + tempMin) / 2).toFixed(1) : '—';

    var soilArr   = (data.hourly && data.hourly.soil_moisture_0_to_1cm) || [];
    var soilValid = soilArr.filter(function(v) { return v != null; });
    var soilMedia = soilValid.length ? (soilValid.reduce(function(a,b) { return a+b; }, 0) / soilValid.length * 100).toFixed(0) : '—';

    var WMO_DESC = { 0:'Céu limpo',1:'Poucas nuvens',2:'Parcialmente nublado',3:'Nublado',45:'Neblina',48:'Névoa',
      51:'Garoa leve',53:'Garoa moderada',55:'Garoa intensa',61:'Chuva leve',63:'Chuva moderada',65:'Chuva intensa',
      80:'Chuva leve',81:'Chuva moderada',82:'Chuva intensa',95:'Tempestade',96:'Tempestade c/ granizo',99:'Tempestade forte' };
    var nextCode = null;
    for (var i2 = 0; i2 < datesArr.length; i2++) {
      if (datesArr[i2] > today) { nextCode = daily.weathercode[i2]; break; }
    }
    var previsao = nextCode != null ? (WMO_DESC[nextCode] || 'Código ' + nextCode) : '—';

    MARKET_DATA.clima = {
      pluvio: pluvio14.toFixed(1) + ' mm',
      umidade: soilMedia !== '—' ? soilMedia + '%' : '—',
      temp: tempMedia !== '—' ? tempMedia + ' °C' : '—',
      previsao: previsao,
      sub: 'Nova Mutum, MT · Open-Meteo'
    };
  } catch(e) {
    MARKET_DATA.clima = { pluvio:'—', umidade:'—', temp:'—', previsao:'—', sub:'Erro ao buscar dados climáticos' };
  }
}

// ── INDICADORES ───────────────────────────────────
function renderIndicadores(emps, lav) {
  var grid = document.getElementById('indicadoresGrid');
  if (!grid) return;

  var totalSaldo = emps.reduce(function(s,e) { return s+e.saldo; }, 0);
  var totalRec   = emps.reduce(function(s,e) { return s+e.recUSD; }, 0);
  var totalPag   = emps.reduce(function(s,e) { return s+e.pagUSD; }, 0);
  var totalExp   = emps.reduce(function(s,e) { return s+e.exposicao; }, 0);

  var sojaRows       = lav.filter(function(l) { return l.cultura === 'SOJA'  && l.precoMedBRL > 0; });
  var precoMedSoja   = sojaRows.length ? sojaRows.reduce(function(s,l) { return s+l.precoMedBRL; }, 0) / sojaRows.length : 0;
  var precoMedSojaUSD = sojaRows.length ? sojaRows.reduce(function(s,l) { return s+l.precoMedUSD; }, 0) / sojaRows.length : 0;
  var milhoRows      = lav.filter(function(l) { return l.cultura === 'MILHO' && l.precoMedBRL > 0; });
  var precoMedMilho  = milhoRows.length ? milhoRows.reduce(function(s,l) { return s+l.precoMedBRL; }, 0) / milhoRows.length : 0;
  var totalProdSoja  = sojaRows.reduce(function(s,l) { return s+l.producao; }, 0);
  var totalVendasSoja = sojaRows.reduce(function(s,l) { return s+l.vendas; }, 0);

  var md = MARKET_DATA;

  var inds = [
    { label: 'Saldo Total (R$)',          value: fBRL(totalSaldo),                         sub: 'Consolidado',                     color: '#40916c', icon: 'fa-wallet' },
    { label: 'Recebimentos USD',           value: 'USD ' + fUSD(totalRec),                  sub: 'Total previsto',                  color: '#2e86de', icon: 'fa-arrow-down' },
    { label: 'Pagamentos USD',             value: 'USD ' + fUSD(totalPag),                  sub: 'Total previsto',                  color: '#e63946', icon: 'fa-arrow-up' },
    { label: 'Exposição Cambial',          value: 'USD ' + fUSD(totalExp),                  sub: totalExp < 0 ? '⚠ Negativa' : 'Positiva', color: totalExp < 0 ? '#e63946' : '#40916c', icon: 'fa-scale-balanced' },
    { label: 'Preço Médio Soja (R$/sc)',   value: 'R$ ' + fNum(precoMedSoja,2) + '/sc',     sub: 'Média ponderada',                 color: '#2d6a4f', icon: 'fa-wheat-awn' },
    { label: 'Preço Médio Soja (USD/sc)',  value: 'USD ' + fNum(precoMedSojaUSD,2) + '/sc', sub: 'Média ponderada',                 color: '#0a9396', icon: 'fa-wheat-awn' },
    { label: 'Preço Médio Milho (R$/sc)',  value: 'R$ ' + fNum(precoMedMilho,2) + '/sc',    sub: 'Média ponderada',                 color: '#e76f51', icon: 'fa-wheat-awn' },
    { label: 'Volume Hedgeado (sc)',        value: fNum(totalVendasSoja),                    sub: 'Soja (safra ativa)',               color: '#40916c', icon: 'fa-chart-bar' },
    { label: 'Volume Exposto (sc)',         value: fNum(totalProdSoja - totalVendasSoja),    sub: 'Soja (safra ativa)',               color: '#e76f51', icon: 'fa-chart-bar' },
    { label: 'Taxa de Câmbio (PTAX)',      value: md.ptax  || '…', sub: md.ptaxSub  || 'Buscando…', color: '#7b2d8b', icon: 'fa-dollar-sign', live: true },
    { label: 'CDI (a.a.)',                 value: md.cdi   || '…', sub: md.cdiSub   || 'Buscando…', color: '#0077b6', icon: 'fa-percent',    live: true },
    { label: 'SELIC (a.a.)',               value: md.selic || '…', sub: md.selicSub || 'Buscando…', color: '#0077b6', icon: 'fa-landmark',   live: true },
    { label: 'Soja Chicago (USc/bu)',      value: md.sojaChicago  || '…', sub: md.sojaChicagoSub  || 'Buscando…', color: '#2d6a4f', icon: 'fa-chart-line', live: true },
    { label: 'Milho Chicago (USc/bu)',     value: md.milhoChicago || '…', sub: md.milhoChicagoSub || 'Buscando…', color: '#e76f51', icon: 'fa-chart-line', live: true },
    { label: 'Pluviometria (14 dias)',      value: md.clima.pluvio,   sub: md.clima.sub,                  color: '#2e86de', icon: 'fa-cloud-rain',      live: true },
    { label: 'Umidade do Solo',            value: md.clima.umidade,  sub: md.clima.sub,                  color: '#0a9396', icon: 'fa-droplet',          live: true },
    { label: 'Temperatura Média',          value: md.clima.temp,     sub: md.clima.sub,                  color: '#e76f51', icon: 'fa-temperature-half', live: true },
    { label: 'Previsão Climática',         value: md.clima.previsao, sub: md.clima.sub + ' · Próx. 24h', color: '#2e86de', icon: 'fa-cloud-sun',        live: true },
  ];

  grid.innerHTML = inds.map(function(ind) {
    return '<div class="ind-card" style="border-left-color:' + ind.color + '">' +
      '<div class="ind-card-top"><div class="ind-label">' + ind.label + '</div>' +
      (ind.live ? '<span class="live-badge"><i class="fas fa-circle"></i> AO VIVO</span>' : '') +
      '</div><div class="ind-value">' + ind.value + '</div>' +
      '<div class="ind-sub">' + ind.sub + '</div></div>';
  }).join('');
}

// ── ALERTAS ───────────────────────────────────────
function renderAlertas(emps) {
  var list = document.getElementById('alertasList');
  if (!list) return;
  var alertas = [];
  emps.forEach(function(e) {
    if      (e.diasRomp <= 7)    alertas.push({ tipo: 'red',    msg: '<strong>' + e.nome + '</strong>: Fluxo rompe em <strong>' + e.diasRomp + ' dia(s)</strong>' });
    else if (e.diasRomp <= 30)   alertas.push({ tipo: 'yellow', msg: '<strong>' + e.nome + '</strong>: Rompimento em ' + e.diasRomp + ' dias' });
    if      (e.exposicao < -1000000) alertas.push({ tipo: 'red',    msg: '<strong>' + e.nome + '</strong>: Exposição cambial negativa elevada (USD ' + fUSD(e.exposicao) + ')' });
    else if (e.exposicao < 0)        alertas.push({ tipo: 'yellow', msg: '<strong>' + e.nome + '</strong>: Exposição cambial negativa (USD ' + fUSD(e.exposicao) + ')' });
    if      (e.saldo < 500000)       alertas.push({ tipo: 'yellow', msg: '<strong>' + e.nome + '</strong>: Saldo disponível crítico (' + fBRL(e.saldo) + ')' });
  });
  if (!alertas.length) alertas.push({ tipo: 'green', msg: 'Nenhum alerta crítico identificado.' });
  list.innerHTML = alertas.map(function(a) {
    return '<div class="alerta alerta-' + a.tipo + '"><div class="alerta-dot dot-' + a.tipo + '"></div><div>' + a.msg + '</div></div>';
  }).join('');
}

// ── Print ──────────────────────────────────────────
var ps = document.createElement('style');
ps.media = 'print';
ps.textContent = '.sidebar,.global-filter-bar{display:none!important}.main-content{margin-left:0!important}.top-header{position:static!important}.hbtn,.btn-upload,.hamburger{display:none!important}';
document.head.appendChild(ps);
