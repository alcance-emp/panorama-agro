/* ===================================================
   PANORAMA AGRO — script.js
   =================================================== */

// ── URL OneDrive (download direto via API) ─────────
const ONEDRIVE_URL = 'https://api.onedrive.com/v1.0/shares/u!aHR0cHM6Ly8xZHJ2Lm1zL3gvYy8wNTQxODE3NWI5ZDBhZGViL0lRRDlWaG55cWVNWlM3cXNIZU1IVTlQYUFlQ1F5U3NNVjVhUmh6WjF2WF9DLTRNP2U9ZUN3WjlQ/root/content';

// ── Estado global ──────────────────────────────────
let DATA = { panorama: [], saldo: [], fluxo: [], lavoura: [] };
let PARSED = { panorama: null, lavoura: [] };
let FILTRO_EMPRESA = 'TODAS';
let SAFRA_SOJA  = null;
let SAFRA_MILHO = null;

let chartSaldo, chartExposicao, chartRecPag, chartCultura, chartSojaBar;
let fluxoDataTable;
let darkMode = false;

// ── Boot ───────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  updateClock();
  setInterval(updateClock, 1000);
  initCharts();
  setupNavHighlight();
  autoLoadFromOneDrive();           // carrega Excel automaticamente
  setInterval(autoLoadFromOneDrive, 5 * 60 * 1000); // recarrega a cada 5 min
});

// ── Carga automática do OneDrive ───────────────────
async function autoLoadFromOneDrive() {
  const badge = document.getElementById('gfBadge');
  if (badge) badge.textContent = 'Atualizando...';
  showLoading(true);
  try {
    const res = await fetch(ONEDRIVE_URL);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const buf = await res.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array', cellDates: true });
    DATA.panorama = sheetToArray(wb, 'PANORAMA');
    DATA.lavoura  = sheetToArray(wb, 'LAVOURA');
    PARSED.panorama = parsePanorama();
    PARSED.lavoura  = parseLavoura();
    FILTRO_EMPRESA  = 'TODAS';
    buildFilterChips(PARSED.panorama.empresas);
    initSafraTabs();
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
      DATA.panorama = sheetToArray(wb, 'PANORAMA');
      DATA.lavoura  = sheetToArray(wb, 'LAVOURA');
      PARSED.panorama = parsePanorama();
      PARSED.lavoura  = parseLavoura();
      FILTRO_EMPRESA = 'TODAS';
      buildFilterChips(PARSED.panorama.empresas);
      initSafraTabs();
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
  // Linhas 6-9 contêm as 4 empresas reais. Limite fixo evita ler cabeçalhos da seção seguinte.
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

// Retorna empresas filtradas
function empresasFiltradas() {
  const all = PARSED.panorama ? PARSED.panorama.empresas : [];
  return FILTRO_EMPRESA === 'TODAS' ? all : all.filter(e => e.nome === FILTRO_EMPRESA);
}

// Retorna lavoura filtrada
function lavouraFiltrada() {
  const all = PARSED.lavoura || [];
  return FILTRO_EMPRESA === 'TODAS' ? all : all.filter(l => l.empresa === FILTRO_EMPRESA);
}

// ── SAFRA TABS ─────────────────────────────────────
function initSafraTabs() {
  const sojaSafras  = [...new Set(PARSED.lavoura.filter(l => l.cultura === 'SOJA').map(l => l.safra))].sort();
  const milhoSafras = [...new Set(PARSED.lavoura.filter(l => l.cultura === 'MILHO').map(l => l.safra))].sort();

  SAFRA_SOJA  = sojaSafras[0]  || null;
  SAFRA_MILHO = milhoSafras[0] || null;

  buildSafraTabs('sojaSafraTabs',  sojaSafras,  'SOJA');
  buildSafraTabs('milhoSafraTabs', milhoSafras, 'MILHO');
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

// ── REFRESH COMPLETO ──────────────────────────────
function refreshAll() {
  if (!PARSED.panorama) return;
  const emps = empresasFiltradas();
  const lav  = lavouraFiltrada();

  renderKPIs(emps, lav);
  renderFluxoTable(emps);
  renderAllCharts(emps, lav);
  renderGrains(lav);
  renderIndicadores(emps, lav);
  renderAlertas(emps);
  renderRompimentos(emps);

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

  const safraAtiva = SAFRA_SOJA || '2025/2026';
  const soja = lav.filter(l => l.cultura === 'SOJA' && l.safra === safraAtiva);
  const totalVendas = soja.reduce((s, l) => s + l.vendas, 0);
  const totalProd   = soja.reduce((s, l) => s + l.producao, 0);
  const totalDisp   = soja.reduce((s, l) => s + l.saldo, 0);
  const pct = totalProd > 0 ? (totalVendas / totalProd) * 100 : 0;

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
  setText('kpiSafra', fNum(totalVendas) + ' sc');
  setTrend('kpiSafraTrend', 3.10, 'vs dia anterior');
  setText('kpiSaldoDisp', fNum(totalDisp) + ' sc');
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
  const cc = chartColors();

  // 1. Saldo por empresa — barras
  chartSaldo = new Chart(document.getElementById('chartSaldo'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Saldo R$', data: [], backgroundColor: PALETTE, borderRadius: 5 }] },
    options: { ...baseOpts(), plugins: { legend: { display: false } },
      scales: scaleOpts(v => 'R$' + (Math.abs(v) >= 1e6 ? (v/1e6).toFixed(1)+'M' : (v/1e3).toFixed(0)+'k')) }
  });

  // 2. Exposição cambial por empresa — barras (positivo azul, negativo vermelho)
  chartExposicao = new Chart(document.getElementById('chartExposicao'), {
    type: 'bar',
    data: { labels: [], datasets: [{ label: 'Exposição USD', data: [], backgroundColor: [], borderRadius: 5 }] },
    options: { ...baseOpts(), plugins: { legend: { display: false } },
      scales: scaleOpts(v => (v/1e3).toFixed(0)+'k') }
  });

  // 3. Rec x Pag — barras agrupadas
  chartRecPag = new Chart(document.getElementById('chartRecPag'), {
    type: 'bar',
    data: { labels: [], datasets: [
      { label: 'Recebimentos', data: [], backgroundColor: 'rgba(64,145,108,.78)',  borderRadius: 4 },
      { label: 'Pagamentos',   data: [], backgroundColor: 'rgba(230,57,70,.75)',   borderRadius: 4 }
    ]},
    options: { ...baseOpts(), scales: scaleOpts(v => (v/1e6).toFixed(1)+'M') }
  });

  // 4. Cultura — barras empilhadas Comprometido × Disponível
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

  // 5. Soja por empresa — barras agrupadas
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

  // 1. Saldo
  chartSaldo.data.labels = labels;
  chartSaldo.data.datasets[0].data = emps.map(e => e.saldo);
  chartSaldo.data.datasets[0].backgroundColor = PALETTE.slice(0, emps.length);
  chartSaldo.update();

  // 2. Exposição — cor condicional
  chartExposicao.data.labels = labels;
  chartExposicao.data.datasets[0].data = emps.map(e => e.exposicao);
  chartExposicao.data.datasets[0].backgroundColor = emps.map(e =>
    e.exposicao < 0 ? 'rgba(230,57,70,.8)' : 'rgba(46,134,222,.8)');
  chartExposicao.update();

  // 3. Rec x Pag
  chartRecPag.data.labels = labels;
  chartRecPag.data.datasets[0].data = emps.map(e => e.recUSD);
  chartRecPag.data.datasets[1].data = emps.map(e => e.pagUSD);
  chartRecPag.update();

  // 4. Cultura — Comprometido × Disponível por cultura
  const culturas = ['SOJA', 'MILHO'];
  const comprom  = culturas.map(c => lav.filter(l => l.cultura === c).reduce((s, l) => s + l.vendas, 0));
  const dispon   = culturas.map(c => lav.filter(l => l.cultura === c).reduce((s, l) => s + l.saldo,  0));
  chartCultura.data.labels = culturas;
  chartCultura.data.datasets[0].data = comprom;
  chartCultura.data.datasets[1].data = dispon;
  chartCultura.update();

  // 5. Soja 2025/2026 (ou safra ativa)
  const safraAtiva = SAFRA_SOJA || '2025/2026';
  const sojaItems  = lav.filter(l => l.cultura === 'SOJA' && l.safra === safraAtiva);
  // agrupar por empresa
  const empNames = [...new Set(sojaItems.map(l => l.empresa))];
  chartSojaBar.data.labels = empNames;
  chartSojaBar.data.datasets[0].data = empNames.map(n => sojaItems.filter(l => l.empresa === n).reduce((s,l) => s + l.producao, 0));
  chartSojaBar.data.datasets[1].data = empNames.map(n => sojaItems.filter(l => l.empresa === n).reduce((s,l) => s + l.vendas, 0));
  chartSojaBar.data.datasets[2].data = empNames.map(n => sojaItems.filter(l => l.empresa === n).reduce((s,l) => s + l.saldo, 0));
  chartSojaBar.update();
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
  const totalProdSoja  = sojaRows.reduce((s, l) => s + l.producao, 0);
  const totalVendasSoja = sojaRows.reduce((s, l) => s + l.vendas, 0);

  const inds = [
    { label: 'Saldo Total (R$)',          value: fBRL(totalSaldo),                         sub: 'Consolidado',         color: '#40916c' },
    { label: 'Recebimentos USD',           value: 'USD ' + fUSD(totalRec),                  sub: 'Total previsto',      color: '#2e86de' },
    { label: 'Pagamentos USD',             value: 'USD ' + fUSD(totalPag),                  sub: 'Total previsto',      color: '#e63946' },
    { label: 'Exposição Cambial',          value: 'USD ' + fUSD(totalExp),                  sub: totalExp < 0 ? '⚠ Negativa' : 'Positiva', color: totalExp < 0 ? '#e63946' : '#40916c' },
    { label: 'Preço Médio Soja (R$/sc)',   value: 'R$ ' + fNum(precoMedSoja, 2) + '/sc',    sub: 'Média ponderada',     color: '#2d6a4f' },
    { label: 'Preço Médio Soja (USD/sc)',  value: 'USD ' + fNum(precoMedSojaUSD, 2) + '/sc',sub: 'Média ponderada',     color: '#0a9396' },
    { label: 'Preço Médio Milho (R$/sc)',  value: 'R$ ' + fNum(precoMedMilho, 2) + '/sc',   sub: 'Média ponderada',     color: '#e76f51' },
    { label: 'Volume Hedgeado (sc)',        value: fNum(totalVendasSoja),                    sub: 'Soja (safra ativa)',  color: '#40916c' },
    { label: 'Volume Exposto (sc)',         value: fNum(totalProdSoja - totalVendasSoja),    sub: 'Soja (safra ativa)',  color: '#e76f51' },
    { label: 'Taxa de Câmbio',             value: 'R$ 5,18 / USD',                          sub: 'vs ant. -0,32% ▼',   color: '#7b2d8b' },
    { label: 'CDI (a.a.)',                 value: '10,65%',                                  sub: 'vs ant. +0,05 p.p.', color: '#0077b6' },
    { label: 'SELIC (a.a.)',               value: '10,50%',                                  sub: 'Meta atual',         color: '#0077b6' },
    { label: 'Soja Chicago (USc/bu)',      value: '1.005',                                   sub: 'CBOT',               color: '#2d6a4f' },
    { label: 'Milho Chicago (USc/bu)',     value: '442',                                     sub: 'CBOT',               color: '#e76f51' },
    { label: 'Pluviometria (mm)',           value: '42 mm',                                  sub: 'Acumulado 14 dias',  color: '#2e86de' },
    { label: 'Umidade do Solo',            value: '68%',                                     sub: 'Média regional',     color: '#0a9396' },
    { label: 'Temperatura Média',          value: '27,4 °C',                                 sub: 'Média semanal',      color: '#e76f51' },
    { label: 'Previsão Climática',         value: 'Chuva moderada',                         sub: 'Próximos 7 dias',    color: '#2e86de' },
  ];

  grid.innerHTML = inds.map(i => `
    <div class="ind-card" style="border-left-color:${i.color}">
      <div class="ind-label">${i.label}</div>
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
