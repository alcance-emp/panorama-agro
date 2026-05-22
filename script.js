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
var DATA = { panorama: [], saldo: [], fluxo: [], lavoura: [], contratos: [] };
var PARSED = { panorama: null, lavoura: [], contratos: [] };
var FILTRO_EMPRESA = 'TODAS';
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
  DATA.panorama  = sheetToArray(wb, 'PANORAMA');
  DATA.lavoura   = sheetToArray(wb, 'LAVOURA');
  DATA.contratos = sheetToArray(wb, 'CONTRATOS');

  PARSED.panorama  = parsePanorama();
  PARSED.lavoura   = parseLavoura();
  PARSED.contratos = parseContratos();

  FILTRO_EMPRESA = 'TODAS';
  buildFilterChips(PARSED.panorama.empresas);
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
var fUSD = function(v) {
  return (v == null || isNaN(v)) ? '—' :
    new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(v);
};
var fNum = function(v, d) {
  d = d || 0;
  return (v == null || isNaN(v)) ? '—' :
    new Intl.NumberFormat('pt-BR', { maximumFractionDigits: d }).format(v);
};
var fPct = function(v) { return (v == null || isNaN(v)) ? '—%' : v.toFixed(1) + '%'; };
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
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r || !r[0]) continue;
    var qtContrato  = parseNum(r[4]);
    var qtEntregue  = parseNum(r[5]);
    var qtAEntregar = parseNum(r[6]);
    var qtFaturada  = parseNum(r[7]);
    var qtAFaturar  = parseNum(r[8]);
    var qtFixada    = parseNum(r[9]);
    var pctAtend    = qtContrato > 0 ? (qtAFaturar / qtContrato) * 100 : 0;
    items.push({
      empresa:     String(r[0] || '').trim(),
      tipo:        String(r[1] || '').trim(),
      cliente:     String(r[2] || '').trim(),
      numContrato: String(r[3] || '').trim(),
      qtContrato:  qtContrato,
      qtEntregue:  qtEntregue,
      qtAEntregar: qtAEntregar,
      qtFaturada:  qtFaturada,
      qtAFaturar:  qtAFaturar,
      qtFixada:    qtFixada,
      cultura:     String(r[10] || '').trim().toUpperCase(),
      safra:       String(r[11] || '').trim(),
      pctAtend:    pctAtend,
    });
  }
  return items;
}

// ── FILTRO GLOBAL DE EMPRESA ──────────────────────
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
  return FILTRO_EMPRESA === 'TODAS' ? all : all.filter(function(l) { return l.empresa === FILTRO_EMPRESA; });
}

function contratosFiltrados() {
  var all = PARSED.contratos || [];
  if (FILTRO_EMPRESA !== 'TODAS') all = all.filter(function(c) { return c.empresa === FILTRO_EMPRESA; });

  var selSafra   = document.getElementById('filtroContratoSafraExtra');
  var safraExtra = selSafra ? selSafra.value : '';
  if (safraExtra) {
    all = all.filter(function(c) { return c.safra === safraExtra; });
  } else if (SAFRA_CONTRATO) {
    all = all.filter(function(c) { return c.safra === SAFRA_CONTRATO; });
  }

  var selCultura   = document.getElementById('filtroContratoCultura');
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
    btn.textContent = s;
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
    if (allSafras.indexOf(c.safra) < 0) allSafras.push(c.safra);
  });
  allSafras.sort();
  SAFRA_CONTRATO = allSafras[0] || '2025/2026';
  buildContratoSafraTabs(allSafras);
}

function buildContratoSafraTabs(safras) {
  var wrap = document.getElementById('contratoSafraTabs');
  if (!wrap) return;
  wrap.innerHTML = '';
  safras.forEach(function(s, i) {
    var btn = document.createElement('button');
    btn.className = 'safra-tab' + (i === 0 ? ' active' : '');
    btn.textContent = s;
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
  renderAllCharts(emps, lav);
  renderGrains(lav);
  renderContratosTable();
  renderIndicadores(emps, lav);
  renderAlertas(emps);

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
  el.innerHTML = arrow + ' ' + Math.abs(pct).toFixed(2) + '% <span style="font-weight:400;opacity:.7">' + label + '</span>';
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

  var headers = ['Empresa','Cultura','Nº Contrato','Cliente / Comprador',
    'Qt. Contrato (sc)','Qt. Entregue (sc)','Qt. A Entregar (sc)',
    'Qt. Faturada (sc)','Qt. A Faturar (sc)','% Atendimento'];

  var rows = items.map(function(c) {
    return [c.empresa, c.cultura, c.numContrato, c.cliente,
      c.qtContrato, c.qtEntregue, c.qtAEntregar, c.qtFaturada, c.qtAFaturar,
      (c.qtContrato > 0 ? (c.qtAFaturar / c.qtContrato * 100).toFixed(1) + '%' : '0%')];
  });

  var wb2 = XLSX.utils.book_new();
  var ws  = XLSX.utils.aoa_to_sheet([headers].concat(rows));
  ws['!cols'] = headers.map(function() { return { wch: 16 }; });
  XLSX.utils.book_append_sheet(wb2, ws, 'Contratos');
  var safra = (document.getElementById('filtroContratoSafraExtra') || {}).value || SAFRA_CONTRATO || 'contratos';
  XLSX.writeFile(wb2, 'contratos_' + safra.replace('/','_') + '.xlsx');
}

function renderContratosTable() {
  var tbody = document.getElementById('contratosBody');
  if (!tbody) return;

  var items = contratosFiltrados();
  tbody.innerHTML = '';

  if (!items.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:var(--text-muted)">Nenhum contrato encontrado</td></tr>';
    return;
  }

  if (AGRUPAR_EMPRESAS) {
    var byEmp = {};
    items.forEach(function(c) { if (!byEmp[c.empresa]) byEmp[c.empresa] = []; byEmp[c.empresa].push(c); });

    Object.keys(byEmp).forEach(function(emp) {
      var contratos = byEmp[emp];
      var totC  = contratos.reduce(function(s,c) { return s+c.qtContrato;  }, 0);
      var totE  = contratos.reduce(function(s,c) { return s+c.qtEntregue;  }, 0);
      var totAE = contratos.reduce(function(s,c) { return s+c.qtAEntregar; }, 0);
      var totF  = contratos.reduce(function(s,c) { return s+c.qtFaturada;  }, 0);
      var totAF = contratos.reduce(function(s,c) { return s+c.qtAFaturar;  }, 0);
      var totPct = totC > 0 ? (totE / totC * 100) : 0;
      var totPctCls = totPct >= 80 ? 'fill-green' : totPct >= 40 ? 'fill-yellow' : 'fill-red';

      var groupRow = document.createElement('tr');
      groupRow.className = 'contrato-group-row';
      groupRow.style.cursor = 'pointer';
      groupRow.innerHTML =
        '<td colspan="4"><span class="group-toggle-icon"><i class="fas fa-chevron-down"></i></span>' +
        '<strong>' + emp + '</strong><span class="group-count">' + contratos.length + ' contrato(s)</span></td>' +
        '<td class="mono"><strong>' + fNum(totC,1) + '</strong></td>' +
        '<td class="mono"><strong>' + fNum(totE,1) + '</strong></td>' +
        '<td class="mono"><strong>' + fNum(totAE,1) + '</strong></td>' +
        '<td class="mono"><strong>' + fNum(totF,1) + '</strong></td>' +
        '<td class="mono"><strong>' + fNum(totAF,1) + '</strong></td>' +
        '<td><div style="display:flex;align-items:center;gap:6px;min-width:90px">' +
        '<div class="progress-bar-wrap" style="flex:1"><div class="progress-bar-fill ' + totPctCls + '" style="width:' + Math.min(totPct,100).toFixed(1) + '%"></div></div>' +
        '<span style="font-size:0.75rem;font-weight:700">' + totPct.toFixed(1) + '%</span></div></td>';
      tbody.appendChild(groupRow);

      var subRows = [];
      contratos.forEach(function(c) {
        var pctAtend  = c.qtContrato > 0 ? (c.qtEntregue / c.qtContrato) * 100 : 0;
        var pctCls    = pctAtend >= 80 ? 'fill-green' : pctAtend >= 40 ? 'fill-yellow' : 'fill-red';
        var pctColor  = pctAtend >= 80 ? 'var(--green-light)' : pctAtend >= 40 ? 'var(--yellow)' : 'var(--red)';
        var culturaCls = c.cultura === 'SOJA' ? 'cultura-soja' : c.cultura === 'MILHO' ? 'cultura-milho' : 'cultura-outro';
        var tr = document.createElement('tr');
        tr.className = 'contrato-sub-row';
        tr.innerHTML =
          '<td style="padding-left:2rem">' + c.empresa + '</td>' +
          '<td><span class="cultura-badge ' + culturaCls + '">' + c.cultura + '</span></td>' +
          '<td class="mono" style="font-size:0.78rem">' + c.numContrato + '</td>' +
          '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + c.cliente + '">' + c.cliente + '</td>' +
          '<td class="mono">' + fNum(c.qtContrato,1) + '</td>' +
          '<td class="mono">' + fNum(c.qtEntregue,1) + '</td>' +
          '<td class="mono">' + fNum(c.qtAEntregar,1) + '</td>' +
          '<td class="mono">' + fNum(c.qtFaturada,1) + '</td>' +
          '<td class="mono">' + fNum(c.qtAFaturar,1) + '</td>' +
          '<td><div style="display:flex;align-items:center;gap:6px;min-width:90px">' +
          '<div class="progress-bar-wrap" style="flex:1"><div class="progress-bar-fill ' + pctCls + '" style="width:' + Math.min(pctAtend,100).toFixed(1) + '%"></div></div>' +
          '<span style="font-size:0.75rem;font-weight:700;color:' + pctColor + '">' + pctAtend.toFixed(1) + '%</span></div></td>';
        tbody.appendChild(tr);
        subRows.push(tr);
      });

      groupRow.addEventListener('click', function() {
        var hidden = subRows[0] && subRows[0].style.display === 'none';
        subRows.forEach(function(r) { r.style.display = hidden ? '' : 'none'; });
        var icon = groupRow.querySelector('.group-toggle-icon i');
        if (icon) icon.className = hidden ? 'fas fa-chevron-down' : 'fas fa-chevron-right';
      });
    });
  } else {
    items.forEach(function(c) {
      var pctAtend  = c.qtContrato > 0 ? (c.qtAFaturar / c.qtContrato) * 100 : 0;
      var pctCls    = pctAtend >= 80 ? 'fill-green' : pctAtend >= 40 ? 'fill-yellow' : 'fill-red';
      var pctColor  = pctAtend >= 80 ? 'var(--green-light)' : pctAtend >= 40 ? 'var(--yellow)' : 'var(--red)';
      var culturaCls = c.cultura === 'SOJA' ? 'cultura-soja' : c.cultura === 'MILHO' ? 'cultura-milho' : 'cultura-outro';
      var tr = document.createElement('tr');
      tr.innerHTML =
        '<td><strong>' + c.empresa + '</strong></td>' +
        '<td><span class="cultura-badge ' + culturaCls + '">' + c.cultura + '</span></td>' +
        '<td class="mono" style="font-size:0.78rem">' + c.numContrato + '</td>' +
        '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + c.cliente + '">' + c.cliente + '</td>' +
        '<td class="mono">' + fNum(c.qtContrato,1) + '</td>' +
        '<td class="mono">' + fNum(c.qtEntregue,1) + '</td>' +
        '<td class="mono">' + fNum(c.qtAEntregar,1) + '</td>' +
        '<td class="mono">' + fNum(c.qtFaturada,1) + '</td>' +
        '<td class="mono">' + fNum(c.qtAFaturar,1) + '</td>' +
        '<td><div style="display:flex;align-items:center;gap:6px;min-width:90px">' +
        '<div class="progress-bar-wrap" style="flex:1"><div class="progress-bar-fill ' + pctCls + '" style="width:' + Math.min(pctAtend,100).toFixed(1) + '%"></div></div>' +
        '<span style="font-size:0.75rem;font-weight:700;color:' + pctColor + '">' + pctAtend.toFixed(1) + '%</span></div></td>';
      tbody.appendChild(tr);
    });
  }

  // Totalizador
  var totalContrato  = items.reduce(function(s,c) { return s+c.qtContrato;  }, 0);
  var totalEntregue  = items.reduce(function(s,c) { return s+c.qtEntregue;  }, 0);
  var totalAEntregar = items.reduce(function(s,c) { return s+c.qtAEntregar; }, 0);
  var totalFaturada  = items.reduce(function(s,c) { return s+c.qtFaturada;  }, 0);
  var totalAFaturar  = items.reduce(function(s,c) { return s+c.qtAFaturar;  }, 0);
  var totalPct       = totalContrato > 0 ? (totalEntregue / totalContrato) * 100 : 0;
  var totalPctCls    = totalPct >= 80 ? 'fill-green' : totalPct >= 40 ? 'fill-yellow' : 'fill-red';

  var tfr = document.createElement('tr');
  tfr.className = 'contrato-total-row';
  tfr.innerHTML =
    '<td colspan="4"><strong>TOTAL (' + items.length + ' contratos)</strong></td>' +
    '<td class="mono"><strong>' + fNum(totalContrato,1) + '</strong></td>' +
    '<td class="mono"><strong>' + fNum(totalEntregue,1) + '</strong></td>' +
    '<td class="mono"><strong>' + fNum(totalAEntregar,1) + '</strong></td>' +
    '<td class="mono"><strong>' + fNum(totalFaturada,1) + '</strong></td>' +
    '<td class="mono"><strong>' + fNum(totalAFaturar,1) + '</strong></td>' +
    '<td><div style="display:flex;align-items:center;gap:6px;min-width:90px">' +
    '<div class="progress-bar-wrap" style="flex:1"><div class="progress-bar-fill ' + totalPctCls + '" style="width:' + Math.min(totalPct,100).toFixed(1) + '%"></div></div>' +
    '<span style="font-size:0.75rem;font-weight:700">' + totalPct.toFixed(1) + '%</span></div></td>';
  tbody.appendChild(tfr);
}

// ── GRÁFICOS ──────────────────────────────────────
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
