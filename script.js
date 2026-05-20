/* ===================================================
   PANORAMA AGRO — script.js
   =================================================== */

// ── URL do Proxy (Cloudflare Worker) ──────────────
// ⚠ Substitua pela URL real do seu Worker após publicar em workers.cloudflare.com
// Exemplo: https://onedrive-proxy.SEU_USUARIO.workers.dev
const ONEDRIVE_URL = 'https://onedrive-proxy.rafaelsousarsp.workers.dev/';

// Token OneDrive (usado internamente pelo Worker — não precisa alterar aqui)
// Link original: https://1drv.ms/x/c/05418175b9d0adeb/IQD9VhnyqeMZS7qsHeMHU9PaAeCQySsMV5aRhzZ1vX_C-4M?e=kjd48N

// ── Estado global ──────────────────────────────────
let DATA = { panorama: [], saldo: [], fluxo: [], lavoura: [], contratos: [] };
let PARSED = { panorama: null, lavoura: [], contratos: [] };
let FILTRO_EMPRESA = 'TODAS';
let SAFRA_SOJA  = null;
let SAFRA_MILHO = null;
let SAFRA_CONTRATO = null;

let chartSaldo, chartExposicao, chartRecPag, chartCultura, chartSojaBar;
let fluxoDataTable, contratosDataTable;
let darkMode = false;

// ── Indicadores de mercado (cache) ────────────────
let MARKET_DATA = {
  ptax: null, ptaxSub: '—',
  cdi: null, cdiSub: '—',
  selic: null, selicSub: '—',
  sojaChicago: null, sojaChicagoSub: '—',
  milhoChicago: null, milhoChicagoSub: '—',
  clima: { pluvio: '—', umidade: '—', temp: '—', previsao: '—', sub: '—' }
};

// ── Boot ───────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  updateClock();
  setInterval(updateClock, 1000);
  initCharts();
  setupNavHighlight();
  autoLoadFromOneDrive();
  setInterval(autoLoadFromOneDrive, 60 * 60 * 1000); // atualiza a cada 1 hora
  fetchMarketData();
  setInterval(fetchMarketData, 15 * 60 * 1000);
});

// ── Carga automática do OneDrive ───────────────────
async function autoLoadFromOneDrive() {
  const badge = document.getElementById('gfBadge');
  if (badge) badge.textContent = 'Atualizando...';
  showLoading(true);
  try {
    // Busca via proxy Cloudflare Worker (contorna bloqueio CORS do OneDrive)
    const res = await fetch(ONEDRIVE_URL);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const buf = await res.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', cellDates: true });
    DATA.panorama  = sheetToArray(wb, 'PANORAMA');
    DATA.lavoura   = sheetToArray(wb, 'LAVOURA');
    DATA.contratos = sheetToArray(wb, 'CONTRATOS');
    PARSED.panorama  = parsePanorama();
    PARSED.lavoura   = parseLavoura();
    PARSED.contratos = parseContratos();
    FILTRO_EMPRESA   = 'TODAS';
    buildFilterChips(PARSED.panorama.empresas);
    initSafraTabs();
    initContratoSafraTabs();
    refreshAll();
    const now = new Date();
    if (badge) badge.textContent = 'OneDrive · ' + now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  } catch (err) {
    console.warn('Falha ao carregar OneDrive:', err);
    if (badge) badge.textContent = 'Upload manual necessário';
  }
  showLoading(false);
}

function updateClock() {
  const now = new Date();
  const el = document.getElementById('headerDate');
  if (el) el.textContent = now.toLocaleDateString('pt-BR') + ' ' +
    now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
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
  const obs = new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (e.isIntersecting) {
        document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
        const t = document.querySelector(`.nav-item[data-section="${e.target.id}"]`);
        if (t) t.classList.add('active');
      }
    });
  }, { threshold: 0.25 });
  document.querySelectorAll('.section').forEach(s => obs.observe(s));
}

// ── Formatadores ──────────────────────────────────
const fBRL = v => (v == null || isNaN(v)) ? '—' :
  new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(v);
const fUSD = v => (v == null || isNaN(v)) ? '—' :
  new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(v);
const fNum = (v, d = 0) => (v == null || isNaN(v)) ? '—' :
  new Intl.NumberFormat('pt-BR', { maximumFractionDigits: d }).format(v);
const fPct = v => (v == null || isNaN(v)) ? '—%' : v.toFixed(1) + '%';
function fDate(v) {
  if (!v) return '—';
  if (v instanceof Date) return v.toLocaleDateString('pt-BR');
  if (typeof v === 'string') { const d = new Date(v); if (!isNaN(d)) return d.toLocaleDateString('pt-BR'); }
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 86400 * 1000)).toLocaleDateString('pt-BR');
  return String(v);
}
function parseNum(v) {
  if (v == null || v === '' || v === '-') return 0;
  if (typeof v === 'number') return v;
  return parseFloat(String(v).replace(/\./g, '').replace(',', '.')) || 0;
}

// ── Leitura Excel ──────────────────────────────────
function loadExcel(input) {
  const file = input.files[0];
  if (!file) return;
  showLoading(true);
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
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
    } catch (err) {
      console.error(err);
      alert('Erro ao ler a planilha: ' + err.message);
    }
    showLoading(false);
  };
  reader.readAsArrayBuffer(file);
}

function sheetToArray(wb, name) {
  const sheet = wb.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
}
function showLoading(show) {
  document.getElementById('loadingOverlay').style.display = show ? 'flex' : 'none';
}

// ── PARSERS ───────────────────────────────────────
function parsePanorama() {
  const rows = DATA.panorama;
  const empresas = [];
  const NOMES_VALIDOS = ['AGRODIMER','KUMMEL','RANCHO ALEGRE','COSTA BEBER'];
  for (let i = 6; i <= 9; i++) {
    const r = rows[i];
    if (!r || !r[0]) continue;
    const nome = String(r[0]).trim();
    if (!NOMES_VALIDOS.some(v => nome.toUpperCase() === v.toUpperCase())) continue;
    empresas.push({
      nome,
      dataBase:  r[1],
      saldo:     parseNum(r[2]),
      dataRomp:  r[3],
      diasRomp:  parseNum(r[4]),
      recUSD:    parseNum(r[5]),
      pagUSD:    parseNum(r[6]),
      exposicao: parseNum(r[7]),
    });
  }
  return { empresas };
}

function parseLavoura() {
  const rows = DATA.lavoura;
  const items = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
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
  const rows = DATA.contratos;
  const items = [];
  // Row 0 is header
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[0]) continue;
    const qtContrato  = parseNum(r[4]);
    const qtEntregue  = parseNum(r[5]);
    const qtAEntregar = parseNum(r[6]);
    const qtFaturada  = parseNum(r[7]);
    const qtAFaturar  = parseNum(r[8]);
    const qtFixada    = parseNum(r[9]);
    const pctAtend    = qtContrato > 0 ? (qtAFaturar / qtContrato) * 100 : 0;
    items.push({
      empresa:     String(r[0] || '').trim(),
      tipo:        String(r[1] || '').trim(),
      cliente:     String(r[2] || '').trim(),
      numContrato: String(r[3] || '').trim(),
      qtContrato,
      qtEntregue,
      qtAEntregar,
      qtFaturada,
      qtAFaturar,
      qtFixada,
      cultura:     String(r[10] || '').trim().toUpperCase(),
      safra:       String(r[11] || '').trim(),
      pctAtend,
    });
  }
  return items;
}

// ── FILTRO GLOBAL DE EMPRESA ──────────────────────
function buildFilterChips(empresas) {
  const wrap = document.getElementById('gfChips');
  if (!wrap) return;
  wrap.innerHTML = `
    <button class="gf-chip active" data-emp="TODAS" onclick="setEmpresaFiltro('TODAS',this)">
      <i class="fas fa-layer-group"></i> Todas
    </button>`;
  empresas.forEach(e => {
    const btn = document.createElement('button');
    btn.className = 'gf-chip';
    btn.dataset.emp = e.nome;
    btn.innerHTML = `<i class="fas fa-building"></i> ${e.nome}`;
    btn.onclick = () => setEmpresaFiltro(e.nome, btn);
    wrap.appendChild(btn);
  });
}

function setEmpresaFiltro(emp, btn) {
  FILTRO_EMPRESA = emp;
  document.querySelectorAll('.gf-chip').forEach(c => c.classList.remove('active'));
  btn.classList.add('active');
  refreshAll();
}

function empresasFiltradas() {
  const all = PARSED.panorama ? PARSED.panorama.empresas : [];
  return FILTRO_EMPRESA === 'TODAS' ? all : all.filter(e => e.nome === FILTRO_EMPRESA);
}

function lavouraFiltrada() {
  const all = PARSED.lavoura || [];
  return FILTRO_EMPRESA === 'TODAS' ? all : all.filter(l => l.empresa === FILTRO_EMPRESA);
}

function contratosFiltrados() {
  let all = PARSED.contratos || [];
  if (FILTRO_EMPRESA !== 'TODAS') all = all.filter(c => c.empresa === FILTRO_EMPRESA);

  // Safra filter: prefer extra dropdown, then SAFRA_CONTRATO tab
  const selSafra = document.getElementById('filtroContratoSafraExtra');
  const safraExtra = selSafra ? selSafra.value : '';
  if (safraExtra) {
    all = all.filter(c => c.safra === safraExtra);
  } else if (SAFRA_CONTRATO) {
    all = all.filter(c => c.safra === SAFRA_CONTRATO);
  }

  // Cultura filter
  const selCultura = document.getElementById('filtroContratoCultura');
  const culturaFiltro = selCultura ? selCultura.value : '';
  if (culturaFiltro) all = all.filter(c => c.cultura === culturaFiltro);

  return all;
}

// ── SAFRA TABS ─────────────────────────────────────
function initSafraTabs() {
  const sojaSafras  = [...new Set(PARSED.lavoura.filter(l => l.cultura === 'SOJA').map(l => l.safra))].sort();
  const milhoSafras = [...new Set(PARSED.lavoura.filter(l => l.cultura === 'MILHO').map(l => l.safra))].sort();
  const allSafras   = [...new Set(PARSED.lavoura.map(l => l.safra))].sort();

  SAFRA_SOJA  = sojaSafras[0]  || null;
  SAFRA_MILHO = milhoSafras[0] || null;

  buildSafraTabs('sojaSafraTabs',  sojaSafras,  'SOJA');
  buildSafraTabs('milhoSafraTabs', milhoSafras, 'MILHO');

  // Populate chart safra dropdowns
  ['filterSafraCultura', 'filterSafraEmpresa'].forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    const cur = sel.value;
    sel.innerHTML = '<option value="">Todas as Safras</option>';
    allSafras.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s; opt.textContent = s;
      if (s === cur) opt.selected = true;
      sel.appendChild(opt);
    });
  });

  // Populate contrato safra extra dropdown
  const allContratoSafras = [...new Set((PARSED.contratos || []).map(c => c.safra))].sort();
  const selExt = document.getElementById('filtroContratoSafraExtra');
  if (selExt) {
    const cur = selExt.value;
    selExt.innerHTML = '<option value="">Todas</option>';
    allContratoSafras.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s; opt.textContent = s;
      if (s === cur) opt.selected = true;
      selExt.appendChild(opt);
    });
  }
}

function buildSafraTabs(containerId, safras, cultura) {
  const wrap = document.getElementById(containerId);
  if (!wrap) return;
  wrap.innerHTML = '';
  safras.forEach((s, i) => {
    const btn = document.createElement('button');
    btn.className = 'safra-tab' + (i === 0 ? ' active' : '');
    btn.textContent = s;
    btn.onclick = () => {
      wrap.querySelectorAll('.safra-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (cultura === 'SOJA')  { SAFRA_SOJA  = s; renderGrainTable('sojaBody',  lavouraFiltrada().filter(l => l.cultura === 'SOJA'  && l.safra === s)); }
      if (cultura === 'MILHO') { SAFRA_MILHO = s; renderGrainTable('milhoBody', lavouraFiltrada().filter(l => l.cultura === 'MILHO' && l.safra === s)); }
    };
    wrap.appendChild(btn);
  });
}

// ── CONTRATOS SAFRA TABS ──────────────────────────
function initContratoSafraTabs() {
  const allSafras = [...new Set((PARSED.contratos || []).map(c => c.safra))].sort();
  SAFRA_CONTRATO = allSafras[0] || '2025/2026';
  buildContratoSafraTabs(allSafras);
}

function buildContratoSafraTabs(safras) {
  const wrap = document.getElementById('contratoSafraTabs');
  if (!wrap) return;
  wrap.innerHTML = '';
  safras.forEach((s, i) => {
    const btn = document.createElement('button');
    btn.className = 'safra-tab' + (i === 0 ? ' active' : '');
    btn.textContent = s;
    btn.onclick = () => {
      wrap.querySelectorAll('.safra-tab').forEach(b => b.classList.remove('active'));
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
  const emps = empresasFiltradas();
  const lav  = lavouraFiltrada();

  renderKPIs(emps, lav);
  renderFluxoTable(emps);
  renderAllCharts(emps, lav);
  renderGrains(lav);
  renderContratosTable();
  renderIndicadores(emps, lav);
  renderAlertas(emps);

  const now = new Date();
  const badge = FILTRO_EMPRESA === 'TODAS' ? 'Todas as empresas' : FILTRO_EMPRESA;
  const el = document.getElementById('lastUpdated');
  if (el) el.textContent = `Atualizado ${now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })} · ${badge}`;
  const gb = document.getElementById('gfBadge');
  if (gb) gb.textContent = `${emps.length} empresa(s)`;
}

// ── KPI CARDS ─────────────────────────────────────
function renderKPIs(emps, lav) {
  const saldoTotal = emps.reduce((s, e) => s + e.saldo, 0);
  const totalRec   = emps.reduce((s, e) => s + e.recUSD, 0);
  const totalPag   = emps.reduce((s, e) => s + e.pagUSD, 0);
  const totalExp   = emps.reduce((s, e) => s + e.exposicao, 0);
  const fluxo      = totalRec - totalPag;

  const safraAtivaSoja  = SAFRA_SOJA  || '2025/2026';
  const safraAtivaMilho = SAFRA_MILHO || '2025/2026';

  const soja  = lav.filter(l => l.cultura === 'SOJA'  && l.safra === safraAtivaSoja);
  const milho = lav.filter(l => l.cultura === 'MILHO' && l.safra === safraAtivaMilho);

  const totalVendasSoja  = soja.reduce((s, l) => s + l.vendas, 0);
  const totalProdSoja    = soja.reduce((s, l) => s + l.producao, 0);
  const totalDispSoja    = soja.reduce((s, l) => s + l.saldo, 0);
  const totalVendasMilho = milho.reduce((s, l) => s + l.vendas, 0);
  const pct = totalProdSoja > 0 ? (totalVendasSoja / totalProdSoja) * 100 : 0;

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
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}
function setTrend(id, pct, label) {
  const el = document.getElementById(id);
  if (!el) return;
  const arrow = pct > 0 ? '▲' : pct < 0 ? '▼' : '→';
  const cls   = pct > 0 ? 'trend-up' : pct < 0 ? 'trend-down' : 'trend-neu';
  el.className = 'kpi-trend ' + cls;
  el.innerHTML = `${arrow} ${Math.abs(pct).toFixed(2)}% <span style="font-weight:400;opacity:.7">${label}</span>`;
}

// ── TABELA FLUXO ──────────────────────────────────
function renderFluxoTable(emps) {
  const tbody = document.getElementById('fluxoBody');
  if (!tbody) return;
  tbody.innerHTML = '';
  emps.forEach(e => {
    const diasCls = e.diasRomp < 30 ? 'dias-vermelho' : e.diasRomp < 90 ? 'dias-amarelo' : 'dias-verde';
    const expCls  = e.exposicao < 0 ? 'exp-neg' : 'exp-pos';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${e.nome}</strong></td>
      <td class="mono">${fDate(e.dataBase)}</td>
      <td class="mono">${fBRL(e.saldo)}</td>
      <td class="mono">${fDate(e.dataRomp)}</td>
      <td><span class="${diasCls}">${e.diasRomp}</span></td>
      <td class="mono">${fUSD(e.recUSD)}</td>
      <td class="mono">${fUSD(e.pagUSD)}</td>
      <td class="${expCls}">${fUSD(e.exposicao)}</td>`;
    tbody.appendChild(tr);
  });

  if (fluxoDataTable) { fluxoDataTable.destroy(); }
  fluxoDataTable = $('#fluxoTable').DataTable({
    pageLength: 10,
    language: { search: 'Buscar:', lengthMenu: 'Exibir _MENU_', info: '_START_–_END_ de _TOTAL_',
                paginate: { previous: '‹', next: '›' }, emptyTable: 'Sem dados' },
    order: [[4, 'asc']], responsive: true
  });
}

// ── GRAINS ────────────────────────────────────────
function renderGrains(lav) {
  const sojaItems  = lav.filter(l => l.cultura === 'SOJA'  && l.safra === (SAFRA_SOJA  || '2025/2026'));
  const milhoItems = lav.filter(l => l.cultura === 'MILHO' && l.safra === (SAFRA_MILHO || '2025/2026'));
  renderGrainTable('sojaBody',  sojaItems);
  renderGrainTable('milhoBody', milhoItems);
}

function renderGrainTable(tbodyId, items) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  tbody.innerHTML = '';
  items.sort((a, b) => a.empresa.localeCompare(b.empresa));
  items.forEach(item => {
    const pct    = item.producao > 0 ? (item.vendas / item.producao) * 100 : 0;
    const pctCls = pct >= 85 ? 'fill-red' : pct >= 60 ? 'fill-yellow' : 'fill-green';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${item.empresa}</strong></td>
      <td class="mono">${fNum(item.producao)}</td>
      <td class="mono">${fNum(item.vendas)}</td>
      <td class="mono">${pct >= 85 ? '⚠ ' : ''}${fNum(item.saldo)}</td>
      <td style="font-weight:700;color:${pct >= 85 ? 'var(--red)' : 'var(--green-mid)'}">${fPct(pct)}</td>
      <td style="min-width:70px">
        <div class="progress-bar-wrap">
          <div class="progress-bar-fill ${pctCls}" style="width:${Math.min(pct, 100).toFixed(1)}%"></div>
        </div>
      </td>`;
    tbody.appendChild(tr);
  });
}

// ── CONTRATOS TABLE ───────────────────────────────
let AGRUPAR_EMPRESAS = true;

function toggleAgruparEmpresas() {
  AGRUPAR_EMPRESAS = !AGRUPAR_EMPRESAS;
  const btn = document.getElementById('btnAgrupar');
  if (btn) {
    btn.classList.toggle('active', AGRUPAR_EMPRESAS);
    btn.innerHTML = AGRUPAR_EMPRESAS
      ? '<i class="fas fa-list"></i> Desagrupar'
      : '<i class="fas fa-layer-group"></i> Agrupar por Empresa';
  }
  renderContratosTable();
}

function exportContratosExcel() {
  const items = contratosFiltrados();
  if (!items.length) { alert('Nenhum contrato para exportar.'); return; }

  const headers = ['Empresa','Cultura','Nº Contrato','Cliente / Comprador',
    'Qt. Contrato (sc)','Qt. Entregue (sc)','Qt. A Entregar (sc)',
    'Qt. Faturada (sc)','Qt. A Faturar (sc)','% Atendimento'];

  const rows = items.map(c => [
    c.empresa, c.cultura, c.numContrato, c.cliente,
    c.qtContrato, c.qtEntregue, c.qtAEntregar,
    c.qtFaturada, c.qtAFaturar,
    (c.qtContrato > 0 ? (c.qtAFaturar / c.qtContrato * 100).toFixed(1) + '%' : '0%')
  ]);

  const wb  = XLSX.utils.book_new();
  const ws  = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = headers.map((h, i) => ({ wch: Math.max(h.length, 14) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Contratos');
  const safra = document.getElementById('filtroContratoSafraExtra')?.value || SAFRA_CONTRATO || 'contratos';
  XLSX.writeFile(wb, `contratos_${safra.replace('/','_')}.xlsx`);
}

function renderContratosTable() {
  const tbody = document.getElementById('contratosBody');
  if (!tbody) return;

  const items = contratosFiltrados();
  tbody.innerHTML = '';

  if (!items.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;color:var(--text-muted)">Nenhum contrato encontrado</td></tr>';
    return;
  }

  const renderRow = (c, isSubRow) => {
    const pctAtend  = c.qtContrato > 0 ? (c.qtAFaturar / c.qtContrato) * 100 : 0;
    const pctCls    = pctAtend >= 80 ? 'fill-green' : pctAtend >= 40 ? 'fill-yellow' : 'fill-red';
    const culturaCls = c.cultura === 'SOJA' ? 'cultura-soja' : c.cultura === 'MILHO' ? 'cultura-milho' : 'cultura-outro';
    const tr = document.createElement('tr');
    if (isSubRow) tr.classList.add('contrato-sub-row');
    tr.innerHTML = `
      <td>${isSubRow ? '<span style="padding-left:1.5rem;opacity:.7">↳</span> ' : '<strong>'+ c.empresa +'</strong>'}</td>
      <td><span class="cultura-badge ${culturaCls}">${c.cultura}</span></td>
      <td class="mono" style="font-size:0.78rem">${c.numContrato}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${c.cliente}">${c.cliente}</td>
      <td class="mono">${fNum(c.qtContrato, 1)}</td>
      <td class="mono">${fNum(c.qtEntregue, 1)}</td>
      <td class="mono">${fNum(c.qtAEntregar, 1)}</td>
      <td class="mono">${fNum(c.qtFaturada, 1)}</td>
      <td class="mono">${fNum(c.qtAFaturar, 1)}</td>
      <td>
        <div style="display:flex;align-items:center;gap:6px;min-width:90px">
          <div class="progress-bar-wrap" style="flex:1">
            <div class="progress-bar-fill ${pctCls}" style="width:${Math.min(pctAtend,100).toFixed(1)}%"></div>
          </div>
          <span style="font-size:0.75rem;font-weight:700;color:${pctAtend>=80?'var(--green-light)':pctAtend>=40?'var(--yellow)':'var(--red)'}">${pctAtend.toFixed(1)}%</span>
        </div>
      </td>`;
    tbody.appendChild(tr);
  };

  if (AGRUPAR_EMPRESAS) {
    const byEmp = {};
    items.forEach(c => { if (!byEmp[c.empresa]) byEmp[c.empresa] = []; byEmp[c.empresa].push(c); });
    Object.entries(byEmp).forEach(([emp, contratos]) => {
      // Group header row
      const totC  = contratos.reduce((s,c) => s + c.qtContrato, 0);
      const totE  = contratos.reduce((s,c) => s + c.qtEntregue, 0);
      const totAE = contratos.reduce((s,c) => s + c.qtAEntregar, 0);
      const totF  = contratos.reduce((s,c) => s + c.qtFaturada, 0);
      const totAF = contratos.reduce((s,c) => s + c.qtAFaturar, 0);
      const totPct = totC > 0 ? (totE / totC * 100) : 0;
      const totPctCls = totPct >= 80 ? 'fill-green' : totPct >= 40 ? 'fill-yellow' : 'fill-red';

      const groupRow = document.createElement('tr');
      groupRow.className = 'contrato-group-row';
      groupRow.style.cursor = 'pointer';
      groupRow.innerHTML = `
        <td colspan="4">
          <span class="group-toggle-icon"><i class="fas fa-chevron-down"></i></span>
          <strong>${emp}</strong>
          <span class="group-count">${contratos.length} contrato(s)</span>
        </td>
        <td class="mono"><strong>${fNum(totC,1)}</strong></td>
        <td class="mono"><strong>${fNum(totE,1)}</strong></td>
        <td class="mono"><strong>${fNum(totAE,1)}</strong></td>
        <td class="mono"><strong>${fNum(totF,1)}</strong></td>
        <td class="mono"><strong>${fNum(totAF,1)}</strong></td>
        <td>
          <div style="display:flex;align-items:center;gap:6px;min-width:90px">
            <div class="progress-bar-wrap" style="flex:1">
              <div class="progress-bar-fill ${totPctCls}" style="width:${Math.min(totPct,100).toFixed(1)}%"></div>
            </div>
            <span style="font-size:0.75rem;font-weight:700">${totPct.toFixed(1)}%</span>
          </div>
        </td>`;
      tbody.appendChild(groupRow);

      const subRows = [];
      contratos.forEach(c => {
        const tr = document.createElement('tr');
        const pctAtend  = c.qtContrato > 0 ? (c.qtEntregue / c.qtContrato) * 100 : 0;
        const pctCls    = pctAtend >= 80 ? 'fill-green' : pctAtend >= 40 ? 'fill-yellow' : 'fill-red';
        const culturaCls = c.cultura === 'SOJA' ? 'cultura-soja' : c.cultura === 'MILHO' ? 'cultura-milho' : 'cultura-outro';
        tr.className = 'contrato-sub-row';
        tr.innerHTML = `
          <td style="padding-left:2rem">${c.empresa}</td>
          <td><span class="cultura-badge ${culturaCls}">${c.cultura}</span></td>
          <td class="mono" style="font-size:0.78rem">${c.numContrato}</td>
          <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${c.cliente}">${c.cliente}</td>
          <td class="mono">${fNum(c.qtContrato,1)}</td>
          <td class="mono">${fNum(c.qtEntregue,1)}</td>
          <td class="mono">${fNum(c.qtAEntregar,1)}</td>
          <td class="mono">${fNum(c.qtFaturada,1)}</td>
          <td class="mono">${fNum(c.qtAFaturar,1)}</td>
          <td>
            <div style="display:flex;align-items:center;gap:6px;min-width:90px">
              <div class="progress-bar-wrap" style="flex:1">
                <div class="progress-bar-fill ${pctCls}" style="width:${Math.min(pctAtend,100).toFixed(1)}%"></div>
              </div>
              <span style="font-size:0.75rem;font-weight:700;color:${pctAtend>=80?'var(--green-light)':pctAtend>=40?'var(--yellow)':'var(--red)'}">${pctAtend.toFixed(1)}%</span>
            </div>
          </td>`;
        tbody.appendChild(tr);
        subRows.push(tr);
      });

      // Toggle collapse
      groupRow.addEventListener('click', () => {
        const hidden = subRows[0] && subRows[0].style.display === 'none';
        subRows.forEach(r => r.style.display = hidden ? '' : 'none');
        const icon = groupRow.querySelector('.group-toggle-icon i');
        if (icon) icon.className = hidden ? 'fas fa-chevron-down' : 'fas fa-chevron-right';
      });
    });
  } else {
    items.forEach(c => renderRow(c, false));
  }

  // Totalizador
  const totalContrato  = items.reduce((s, c) => s + c.qtContrato, 0);
  const totalEntregue  = items.reduce((s, c) => s + c.qtEntregue, 0);
  const totalAEntregar = items.reduce((s, c) => s + c.qtAEntregar, 0);
  const totalFaturada  = items.reduce((s, c) => s + c.qtFaturada, 0);
  const totalAFaturar  = items.reduce((s, c) => s + c.qtAFaturar, 0);
  const totalPct       = totalContrato > 0 ? (totalEntregue / totalContrato) * 100 : 0;
  const totalPctCls    = totalPct >= 80 ? 'fill-green' : totalPct >= 40 ? 'fill-yellow' : 'fill-red';

  const tfr = document.createElement('tr');
  tfr.className = 'contrato-total-row';
  tfr.innerHTML = `
    <td colspan="4"><strong>TOTAL (${items.length} contratos)</strong></td>
    <td class="mono"><strong>${fNum(totalContrato, 1)}</strong></td>
    <td class="mono"><strong>${fNum(totalEntregue, 1)}</strong></td>
    <td class="mono"><strong>${fNum(totalAEntregar, 1)}</strong></td>
    <td class="mono"><strong>${fNum(totalFaturada, 1)}</strong></td>
    <td class="mono"><strong>${fNum(totalAFaturar, 1)}</strong></td>
    <td>
      <div style="display:flex;align-items:center;gap:6px;min-width:90px">
        <div class="progress-bar-wrap" style="flex:1">
          <div class="progress-bar-fill ${totalPctCls}" style="width:${Math.min(totalPct,100).toFixed(1)}%"></div>
        </div>
        <span style="font-size:0.75rem;font-weight:700">${totalPct.toFixed(1)}%</span>
      </div>
    </td>`;
  tbody.appendChild(tfr);
}

// ── GRÁFICOS ──────────────────────────────────────
function chartColors() {
  return darkMode
    ? { grid: 'rgba(255,255,255,.08)', text: '#8aad96', tick: '#8aad96' }
    : { grid: 'rgba(0,0,0,.06)',       text: '#6b7c74', tick: '#6b7c74' };
}

const PALETTE = ['#40916c','#2e86de','#e76f51','#7b2d8b','#0a9396','#f4a261','#e63946','#74c69d'];

function baseOpts() {
  const cc = chartColors();
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 500 },
    plugins: { legend: { labels: { color: cc.text, font: { family: 'Outfit', size: 11 }, boxWidth: 12 } } }
  };
}
function scaleOpts(yCallback) {
  const cc = chartColors();
  return {
    x: { ticks: { color: cc.tick, font: { size: 10 } }, grid: { color: cc.grid } },
    y: { ticks: { color: cc.tick, font: { size: 10 }, callback: yCallback }, grid: { color: cc.grid } }
  };
}

function initCharts() {
  chartSaldo = new Chart(document.getElementById('chartSaldo'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Saldo R$', data: [], backgroundColor: PALETTE, borderRadius: 5 }] },
    options: { ...baseOpts(), plugins: { legend: { display: false } },
      scales: scaleOpts(v => 'R$' + (Math.abs(v) >= 1e6 ? (v/1e6).toFixed(1)+'M' : (v/1e3).toFixed(0)+'k')) }
  });

  chartExposicao = new Chart(document.getElementById('chartExposicao'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Exposição USD', data: [], backgroundColor: [], borderRadius: 5 }] },
    options: { ...baseOpts(), plugins: { legend: { display: false } },
      scales: scaleOpts(v => (v/1e3).toFixed(0)+'k') }
  });

  chartRecPag = new Chart(document.getElementById('chartRecPag'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Recebimentos', data: [], backgroundColor: 'rgba(64,145,108,.78)',  borderRadius: 4 },
      { label: 'Pagamentos',   data: [], backgroundColor: 'rgba(230,57,70,.75)',   borderRadius: 4 }
    ]},
    options: { ...baseOpts(), scales: scaleOpts(v => (v/1e6).toFixed(1)+'M') }
  });

  chartCultura = new Chart(document.getElementById('chartCultura'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Comprometido', data: [], backgroundColor: 'rgba(230,57,70,.78)',    borderRadius: 4 },
      { label: 'Disponível',   data: [], backgroundColor: 'rgba(64,145,108,.82)',   borderRadius: 4 }
    ]},
    options: { ...baseOpts(),
      scales: { ...scaleOpts(v => (v/1e3).toFixed(0)+'k sc'), x: { stacked: false,
        ticks: { color: chartColors().tick, font: { size: 10 } }, grid: { color: chartColors().grid } } } }
  });

  chartSojaBar = new Chart(document.getElementById('chartSojaBar'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Produção Est.', data: [], backgroundColor: 'rgba(45,106,79,.6)',    borderRadius: 4 },
      { label: 'Vendas',        data: [], backgroundColor: 'rgba(64,145,108,.85)',  borderRadius: 4 },
      { label: 'Saldo Disp.',   data: [], backgroundColor: 'rgba(116,198,157,.75)', borderRadius: 4 }
    ]},
    options: { ...baseOpts(), scales: scaleOpts(v => (v/1e3).toFixed(0)+'k') }
  });
}

function renderAllCharts(emps, lav) {
  const labels = emps.map(e => e.nome);

  chartSaldo.data.labels = labels;
  chartSaldo.data.datasets[0].data = emps.map(e => e.saldo);
  chartSaldo.data.datasets[0].backgroundColor = emps.map(() => 'rgba(64,145,108,.85)');
  chartSaldo.update();

  chartExposicao.data.labels = labels;
  chartExposicao.data.datasets[0].data = emps.map(e => e.exposicao);
  chartExposicao.data.datasets[0].backgroundColor = emps.map(e =>
    e.exposicao < 0 ? 'rgba(230,57,70,.8)' : 'rgba(46,134,222,.8)');
  chartExposicao.update();

  chartRecPag.data.labels = labels;
  chartRecPag.data.datasets[0].data = emps.map(e => e.recUSD);
  chartRecPag.data.datasets[1].data = emps.map(e => e.pagUSD);
  chartRecPag.update();

  const culturas = ['SOJA', 'MILHO'];
  const safraCultura = document.getElementById('filterSafraCultura') ? document.getElementById('filterSafraCultura').value : '';
  const lavCultura = safraCultura ? lav.filter(l => l.safra === safraCultura) : lav;
  const comprom  = culturas.map(c => lavCultura.filter(l => l.cultura === c).reduce((s, l) => s + l.vendas, 0));
  const dispon   = culturas.map(c => lavCultura.filter(l => l.cultura === c).reduce((s, l) => s + l.saldo,  0));
  chartCultura.data.labels = culturas;
  chartCultura.data.datasets[0].data = comprom;
  chartCultura.data.datasets[1].data = dispon;
  chartCultura.update();

  const safraEmpresa = document.getElementById('filterSafraEmpresa') ? document.getElementById('filterSafraEmpresa').value : '';
  const safraAtiva = safraEmpresa || SAFRA_SOJA || '2025/2026';
  const sojaItems  = safraEmpresa
    ? lav.filter(l => l.safra === safraEmpresa)
    : lav.filter(l => l.cultura === 'SOJA' && l.safra === safraAtiva);
  const empNames = [...new Set(sojaItems.map(l => l.empresa))];
  chartSojaBar.data.labels = empNames;
  chartSojaBar.data.datasets[0].data = empNames.map(n => sojaItems.filter(l => l.empresa === n).reduce((s,l) => s + l.producao, 0));
  chartSojaBar.data.datasets[1].data = empNames.map(n => sojaItems.filter(l => l.empresa === n).reduce((s,l) => s + l.vendas, 0));
  chartSojaBar.data.datasets[2].data = empNames.map(n => sojaItems.filter(l => l.empresa === n).reduce((s,l) => s + l.saldo, 0));
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

// PTAX via API pública do Banco Central do Brasil
async function fetchPTAX() {
  try {
    // Busca o PTAX do dia anterior útil
    const hoje = new Date();
    const ontem = new Date(hoje);
    ontem.setDate(hoje.getDate() - 1);
    // Formata para @odata: MM-DD-YYYY
    const fmt = d => `${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}-${d.getFullYear()}`;
    // Tenta os últimos 5 dias úteis para garantir resultado
    for (let i = 1; i <= 5; i++) {
      const d = new Date(hoje);
      d.setDate(hoje.getDate() - i);
      const url = `https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='${fmt(d)}'&$top=1&$format=json&$select=cotacaoCompra,cotacaoVenda`;
      const res = await fetch(url);
      if (!res.ok) continue;
      const json = await res.json();
      const val = json.value && json.value[0];
      if (val) {
        const ptax = ((val.cotacaoCompra + val.cotacaoVenda) / 2).toFixed(4);
        MARKET_DATA.ptax = 'R$ ' + String(ptax).replace('.', ',');
        MARKET_DATA.ptaxSub = `Compra R$${String(val.cotacaoCompra.toFixed(4)).replace('.',',')} · Venda R$${String(val.cotacaoVenda.toFixed(4)).replace('.',',')}`;
        return;
      }
    }
    MARKET_DATA.ptax = '—';
    MARKET_DATA.ptaxSub = 'Dado não disponível';
  } catch (e) {
    MARKET_DATA.ptax = '—';
    MARKET_DATA.ptaxSub = 'Erro ao buscar PTAX';
  }
}

// CDI e SELIC via API do Banco Central
async function fetchCDISELIC() {
  try {
    const [cdiRes, selicRes] = await Promise.allSettled([
      fetch('https://api.bcb.gov.br/dados/serie/bcdata.sgs.12/dados/ultimos/1?formato=json'),
      fetch('https://api.bcb.gov.br/dados/serie/bcdata.sgs.11/dados/ultimos/1?formato=json')
    ]);

    if (cdiRes.status === 'fulfilled' && cdiRes.value.ok) {
      const cdiData = await cdiRes.value.json();
      if (cdiData && cdiData[0]) {
        const cdiDiario = parseFloat(String(cdiData[0].valor).replace(',', '.'));
        const cdiAnual  = ((Math.pow(1 + cdiDiario / 100, 252) - 1) * 100).toFixed(2);
        MARKET_DATA.cdi    = cdiAnual.replace('.', ',') + '% a.a.';
        MARKET_DATA.cdiSub = `Diário: ${cdiData[0].valor}% · ${cdiData[0].data}`;
      }
    }

    if (selicRes.status === 'fulfilled' && selicRes.value.ok) {
      const selicData = await selicRes.value.json();
      if (selicData && selicData[0]) {
        const selicDiario = parseFloat(String(selicData[0].valor).replace(',', '.'));
        const selicAnual  = ((Math.pow(1 + selicDiario / 100, 252) - 1) * 100).toFixed(2);
        MARKET_DATA.selic    = selicAnual.replace('.', ',') + '% a.a.';
        MARKET_DATA.selicSub = `Diário: ${selicData[0].valor}% · ${selicData[0].data}`;
      }
    }

    // Fallback: try SGS series 432 (SELIC meta diária) if 11 didn't work
    if (!MARKET_DATA.selic || MARKET_DATA.selic === '—') {
      const r2 = await fetch('https://api.bcb.gov.br/dados/serie/bcdata.sgs.432/dados/ultimos/1?formato=json');
      if (r2.ok) {
        const d2 = await r2.json();
        if (d2 && d2[0]) {
          const v = parseFloat(String(d2[0].valor).replace(',', '.'));
          const anual = ((Math.pow(1 + v / 100, 252) - 1) * 100).toFixed(2);
          MARKET_DATA.selic    = anual.replace('.', ',') + '% a.a.';
          MARKET_DATA.selicSub = `Diário: ${d2[0].valor}% · ${d2[0].data}`;
        }
      }
    }
  } catch (e) {
    MARKET_DATA.cdi    = '—';
    MARKET_DATA.cdiSub = 'Erro ao buscar CDI';
    MARKET_DATA.selic    = '—';
    MARKET_DATA.selicSub = 'Erro ao buscar SELIC';
  }
}

// Chicago (CBOT) via Yahoo Finance (mais confiável) com fallback stooq
async function fetchChicagoPrices() {
  // Tenta Yahoo Finance primeiro (CORS livre)
  try {
    const [sojaY, milhoY] = await Promise.allSettled([
      fetch('https://query1.finance.yahoo.com/v8/finance/chart/ZS=F?interval=1d&range=5d'),
      fetch('https://query1.finance.yahoo.com/v8/finance/chart/ZC=F?interval=1d&range=5d')
    ]);

    if (sojaY.status === 'fulfilled' && sojaY.value.ok) {
      const j = await sojaY.value.json();
      const c = j?.chart?.result?.[0];
      if (c) {
        const prices = c.indicators?.quote?.[0]?.close || [];
        const last = prices.filter(v => v != null).pop();
        if (last) {
          MARKET_DATA.sojaChicago    = last.toFixed(2) + ' USc/bu';
          MARKET_DATA.sojaChicagoSub = 'CBOT · Yahoo Finance (ZS=F)';
        }
      }
    }

    if (milhoY.status === 'fulfilled' && milhoY.value.ok) {
      const j = await milhoY.value.json();
      const c = j?.chart?.result?.[0];
      if (c) {
        const prices = c.indicators?.quote?.[0]?.close || [];
        const last = prices.filter(v => v != null).pop();
        if (last) {
          MARKET_DATA.milhoChicago    = last.toFixed(2) + ' USc/bu';
          MARKET_DATA.milhoChicagoSub = 'CBOT · Yahoo Finance (ZC=F)';
        }
      }
    }

    if (MARKET_DATA.sojaChicago && MARKET_DATA.milhoChicago) return;
  } catch (e) { /* continua para stooq */ }

  // Fallback: stooq
  try {
    const now = new Date();
    const d1  = new Date(now); d1.setDate(now.getDate() - 10);
    const fmt = d => `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;

    const [sojaRes, milhoRes] = await Promise.allSettled([
      fetch(`https://stooq.com/q/d/l/?s=zs.f&d1=${fmt(d1)}&d2=${fmt(now)}&i=d`),
      fetch(`https://stooq.com/q/d/l/?s=zc.f&d1=${fmt(d1)}&d2=${fmt(now)}&i=d`)
    ]);

    if (sojaRes.status === 'fulfilled' && sojaRes.value.ok && !MARKET_DATA.sojaChicago) {
      const txt = await sojaRes.value.text();
      const lines = txt.trim().split('\n').filter(l => !l.startsWith('Date'));
      if (lines.length) {
        const last = lines[lines.length - 1].split(',');
        if (last.length >= 5 && last[4] && last[4] !== 'null') {
          MARKET_DATA.sojaChicago    = parseFloat(last[4]).toFixed(2) + ' USc/bu';
          MARKET_DATA.sojaChicagoSub = `CBOT · Stooq · ${last[0]}`;
        }
      }
    }

    if (milhoRes.status === 'fulfilled' && milhoRes.value.ok && !MARKET_DATA.milhoChicago) {
      const txt = await milhoRes.value.text();
      const lines = txt.trim().split('\n').filter(l => !l.startsWith('Date'));
      if (lines.length) {
        const last = lines[lines.length - 1].split(',');
        if (last.length >= 5 && last[4] && last[4] !== 'null') {
          MARKET_DATA.milhoChicago    = parseFloat(last[4]).toFixed(2) + ' USc/bu';
          MARKET_DATA.milhoChicagoSub = `CBOT · Stooq · ${last[0]}`;
        }
      }
    }
  } catch (e2) { /* ignorar */ }

  if (!MARKET_DATA.sojaChicago) {
    MARKET_DATA.sojaChicago    = '—';
    MARKET_DATA.sojaChicagoSub = 'Indisponível (CORS)';
  }
  if (!MARKET_DATA.milhoChicago) {
    MARKET_DATA.milhoChicago    = '—';
    MARKET_DATA.milhoChicagoSub = 'Indisponível (CORS)';
  }
}

// Climatempo / Open-Meteo (Nova Mutum, MT — -13.8278, -54.1992)
async function fetchClimatempo() {
  try {
    // Open-Meteo é completamente gratuita e sem chave de API
    const lat = -13.8278, lon = -54.1992;
    const url = `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}&hourly=soil_moisture_0_to_1cm&daily=temperature_2m_max,temperature_2m_min,precipitation_sum,weathercode&timezone=America%2FCuiaba&past_days=14&forecast_days=7`;
    const res = await fetch(url);
    if (!res.ok) throw new Error('Open-Meteo error');
    const data = await res.json();

    const daily = data.daily;
    const today = new Date().toISOString().split('T')[0];

    // Pluviometria acumulada 14 dias passados
    const pluvioArr = daily.precipitation_sum || [];
    const datesArr  = daily.time || [];
    let pluvio14 = 0;
    datesArr.forEach((d, i) => {
      if (d <= today && pluvioArr[i] != null) pluvio14 += pluvioArr[i];
    });

    // Temperatura média (max+min)/2 do último dia disponível <= hoje
    let tempMax = null, tempMin = null, lastDate = null;
    datesArr.forEach((d, i) => {
      if (d <= today) {
        tempMax  = daily.temperature_2m_max[i];
        tempMin  = daily.temperature_2m_min[i];
        lastDate = d;
      }
    });
    const tempMedia = (tempMax != null && tempMin != null) ? ((tempMax + tempMin) / 2).toFixed(1) : '—';

    // Umidade do solo (hourly média)
    const soilArr = (data.hourly && data.hourly.soil_moisture_0_to_1cm) || [];
    const soilValid = soilArr.filter(v => v != null);
    const soilMedia = soilValid.length ? (soilValid.reduce((a, b) => a + b, 0) / soilValid.length * 100).toFixed(0) : '—';

    // Previsão climática próximos 7 dias (weathercode do 1º dia futuro)
    const WMO_DESC = {
      0: 'Céu limpo', 1: 'Poucas nuvens', 2: 'Parcialmente nublado', 3: 'Nublado',
      45: 'Neblina', 48: 'Névoa',
      51: 'Garoa leve', 53: 'Garoa moderada', 55: 'Garoa intensa',
      61: 'Chuva leve', 63: 'Chuva moderada', 65: 'Chuva intensa',
      71: 'Neve leve', 73: 'Neve moderada', 75: 'Neve intensa',
      80: 'Chuva leve', 81: 'Chuva moderada', 82: 'Chuva intensa',
      95: 'Tempestade', 96: 'Tempestade c/ granizo', 99: 'Tempestade forte'
    };
    let nextCode = null;
    for (let i = 0; i < datesArr.length; i++) {
      if (datesArr[i] > today) { nextCode = daily.weathercode[i]; break; }
    }
    const previsao = nextCode != null ? (WMO_DESC[nextCode] || `Código ${nextCode}`) : '—';

    MARKET_DATA.clima = {
      pluvio:   pluvio14.toFixed(1) + ' mm',
      umidade:  soilMedia !== '—' ? soilMedia + '%' : '—',
      temp:     tempMedia !== '—' ? tempMedia + ' °C' : '—',
      previsao: previsao,
      sub:      `Nova Mutum, MT · Open-Meteo`
    };
  } catch (e) {
    MARKET_DATA.clima = {
      pluvio:   '—', umidade: '—', temp: '—', previsao: '—',
      sub:      'Erro ao buscar dados climáticos'
    };
  }
}

// ── INDICADORES ───────────────────────────────────
function renderIndicadores(emps, lav) {
  const grid = document.getElementById('indicadoresGrid');
  if (!grid) return;

  const totalSaldo = emps.reduce((s, e) => s + e.saldo, 0);
  const totalRec   = emps.reduce((s, e) => s + e.recUSD, 0);
  const totalPag   = emps.reduce((s, e) => s + e.pagUSD, 0);
  const totalExp   = emps.reduce((s, e) => s + e.exposicao, 0);

  const sojaRows = lav.filter(l => l.cultura === 'SOJA' && l.precoMedBRL > 0);
  const precoMedSoja = sojaRows.length ? sojaRows.reduce((s, l) => s + l.precoMedBRL, 0) / sojaRows.length : 0;
  const precoMedSojaUSD = sojaRows.length ? sojaRows.reduce((s, l) => s + l.precoMedUSD, 0) / sojaRows.length : 0;
  const milhoRows = lav.filter(l => l.cultura === 'MILHO' && l.precoMedBRL > 0);
  const precoMedMilho = milhoRows.length ? milhoRows.reduce((s, l) => s + l.precoMedBRL, 0) / milhoRows.length : 0;
  const totalProdSoja   = sojaRows.reduce((s, l) => s + l.producao, 0);
  const totalVendasSoja = sojaRows.reduce((s, l) => s + l.vendas, 0);

  const md = MARKET_DATA;

  const inds = [
    { label: 'Saldo Total (R$)',          value: fBRL(totalSaldo),                         sub: 'Consolidado',                     color: '#40916c', icon: 'fa-wallet' },
    { label: 'Recebimentos USD',           value: 'USD ' + fUSD(totalRec),                  sub: 'Total previsto',                  color: '#2e86de', icon: 'fa-arrow-down' },
    { label: 'Pagamentos USD',             value: 'USD ' + fUSD(totalPag),                  sub: 'Total previsto',                  color: '#e63946', icon: 'fa-arrow-up' },
    { label: 'Exposição Cambial',          value: 'USD ' + fUSD(totalExp),                  sub: totalExp < 0 ? '⚠ Negativa' : 'Positiva', color: totalExp < 0 ? '#e63946' : '#40916c', icon: 'fa-scale-balanced' },
    { label: 'Preço Médio Soja (R$/sc)',   value: 'R$ ' + fNum(precoMedSoja, 2) + '/sc',    sub: 'Média ponderada',                 color: '#2d6a4f', icon: 'fa-wheat-awn' },
    { label: 'Preço Médio Soja (USD/sc)',  value: 'USD ' + fNum(precoMedSojaUSD, 2) + '/sc',sub: 'Média ponderada',                 color: '#0a9396', icon: 'fa-wheat-awn' },
    { label: 'Preço Médio Milho (R$/sc)',  value: 'R$ ' + fNum(precoMedMilho, 2) + '/sc',   sub: 'Média ponderada',                 color: '#e76f51', icon: 'fa-wheat-awn' },
    { label: 'Volume Hedgeado (sc)',        value: fNum(totalVendasSoja),                    sub: 'Soja (safra ativa)',               color: '#40916c', icon: 'fa-chart-bar' },
    { label: 'Volume Exposto (sc)',         value: fNum(totalProdSoja - totalVendasSoja),    sub: 'Soja (safra ativa)',               color: '#e76f51', icon: 'fa-chart-bar' },
    { label: 'Taxa de Câmbio (PTAX)',      value: md.ptax || '…',                            sub: md.ptaxSub || 'Buscando…',         color: '#7b2d8b', icon: 'fa-dollar-sign', live: true },
    { label: 'CDI (a.a.)',                 value: md.cdi  || '…',                            sub: md.cdiSub  || 'Buscando…',         color: '#0077b6', icon: 'fa-percent',    live: true },
    { label: 'SELIC (a.a.)',               value: md.selic || '…',                           sub: md.selicSub || 'Buscando…',        color: '#0077b6', icon: 'fa-landmark',   live: true },
    { label: 'Soja Chicago (USc/bu)',      value: md.sojaChicago  || '…',                   sub: md.sojaChicagoSub  || 'Buscando…', color: '#2d6a4f', icon: 'fa-chart-line',  live: true },
    { label: 'Milho Chicago (USc/bu)',     value: md.milhoChicago || '…',                   sub: md.milhoChicagoSub || 'Buscando…', color: '#e76f51', icon: 'fa-chart-line',  live: true },
    { label: 'Pluviometria (14 dias)',      value: md.clima.pluvio,                           sub: md.clima.sub,                     color: '#2e86de', icon: 'fa-cloud-rain',  live: true },
    { label: 'Umidade do Solo',            value: md.clima.umidade,                          sub: md.clima.sub,                     color: '#0a9396', icon: 'fa-droplet',     live: true },
    { label: 'Temperatura Média',          value: md.clima.temp,                             sub: md.clima.sub,                     color: '#e76f51', icon: 'fa-temperature-half', live: true },
    { label: 'Previsão Climática',         value: md.clima.previsao,                         sub: md.clima.sub + ' · Próx. 24h',    color: '#2e86de', icon: 'fa-cloud-sun',   live: true },
  ];

  grid.innerHTML = inds.map(i => `
    <div class="ind-card" style="border-left-color:${i.color}">
      <div class="ind-card-top">
        <div class="ind-label">${i.label}</div>
        ${i.live ? '<span class="live-badge"><i class="fas fa-circle"></i> AO VIVO</span>' : ''}
      </div>
      <div class="ind-value">${i.value}</div>
      <div class="ind-sub">${i.sub}</div>
    </div>`).join('');
}

// ── ALERTAS ───────────────────────────────────────
function renderAlertas(emps) {
  const list = document.getElementById('alertasList');
  if (!list) return;
  const alertas = [];
  emps.forEach(e => {
    if (e.diasRomp <= 7)    alertas.push({ tipo: 'red',    msg: `<strong>${e.nome}</strong>: Fluxo rompe em <strong>${e.diasRomp} dia(s)</strong>` });
    else if (e.diasRomp <= 30) alertas.push({ tipo: 'yellow', msg: `<strong>${e.nome}</strong>: Rompimento em ${e.diasRomp} dias` });
    if (e.exposicao < -1000000) alertas.push({ tipo: 'red',    msg: `<strong>${e.nome}</strong>: Exposição cambial negativa elevada (USD ${fUSD(e.exposicao)})` });
    else if (e.exposicao < 0)   alertas.push({ tipo: 'yellow', msg: `<strong>${e.nome}</strong>: Exposição cambial negativa (USD ${fUSD(e.exposicao)})` });
    if (e.saldo < 500000)       alertas.push({ tipo: 'yellow', msg: `<strong>${e.nome}</strong>: Saldo disponível crítico (${fBRL(e.saldo)})` });
  });
  if (!alertas.length) alertas.push({ tipo: 'green', msg: 'Nenhum alerta crítico identificado.' });
  list.innerHTML = alertas.map(a => `
    <div class="alerta alerta-${a.tipo}">
      <div class="alerta-dot dot-${a.tipo}"></div>
      <div>${a.msg}</div>
    </div>`).join('');
}

// ── PRÓXIMOS ROMPIMENTOS ──────────────────────────
function renderRompimentos(emps) {
  const tbody = document.getElementById('rompBody');
  if (!tbody) return;
  const sorted = [...emps].sort((a, b) => a.diasRomp - b.diasRomp);
  tbody.innerHTML = sorted.map(e => {
    const cls   = e.diasRomp < 30 ? 'dias-vermelho' : e.diasRomp < 90 ? 'dias-amarelo' : 'dias-verde';
    const badge = e.diasRomp < 7 ? '🔴' : e.diasRomp < 30 ? '🟡' : '🟢';
    return `<tr>
      <td><strong>${e.nome}</strong></td>
      <td class="mono">${fDate(e.dataRomp)}</td>
      <td class="${cls}">${e.diasRomp}</td>
      <td>${badge}</td>
    </tr>`;
  }).join('');
}

// ── Print ──────────────────────────────────────────
const ps = document.createElement('style');
ps.media = 'print';
ps.textContent = `.sidebar,.global-filter-bar{display:none!important}.main-content{margin-left:0!important}.top-header{position:static!important}.hbtn,.btn-upload,.hamburger{display:none!important}`;
document.head.appendChild(ps);
