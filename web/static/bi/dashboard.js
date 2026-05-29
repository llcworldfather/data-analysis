const TOKEN = window.BI_TOKEN;
const CAPABILITIES = window.BI_CAPABILITIES;


/* ── State ── */
let allData       = [];   // full dataset from /api/bi/summary
let filteredData  = [];   // after main filters
let sortCol       = "全局总成本贡献";
let sortAsc       = false;
let detailRaw     = [];   // full detail from /api/bi/detail
let detailFiltered = [];  // after detail filter
let hasName       = false;
let hasUnit       = false;
let maxPct        = 0;

/* ── Init ── */
function applyAiCapabilities() {
  const ai = (CAPABILITIES && CAPABILITIES.ai_purchase_advice) || {};
  const btn = document.getElementById('ai-advice-btn');
  const disabledPanel = document.getElementById('ai-advice-disabled');
  const baiduHint = document.getElementById('ai-advice-baidu-hint');

  if (!ai.enabled) {
    if (btn) {
      btn.disabled = true;
      btn.title = ai.reason || 'AI 未配置';
    }
    if (disabledPanel) {
      document.getElementById('ai-advice-disabled-reason').textContent =
        ai.reason || '未配置 DeepSeek API Key';
      show('ai-advice-disabled', true);
    }
    show('ai-advice-baidu-hint', false);
    return;
  }

  if (btn) {
    btn.disabled = false;
    btn.title = '';
  }
  show('ai-advice-disabled', false);
  if (ai.baidu_note && baiduHint) {
    document.getElementById('ai-advice-baidu-hint-text').textContent = ai.baidu_note;
    show('ai-advice-baidu-hint', true);
  } else {
    show('ai-advice-baidu-hint', false);
  }
}

window.addEventListener('DOMContentLoaded', () => {
  applyAiCapabilities();
  loadData();

  // 事件委托：点击任意 .count-badge 按钮时读取 data-comp 属性打开详情
  // 避免 inline onclick 中因双引号冲突导致的 SyntaxError
  document.getElementById('main-tbody').addEventListener('click', function(e) {
    const btn = e.target.closest('.count-badge');
    if (btn) showDetail(btn.dataset.comp);
  });

  document.getElementById('detail-modal').addEventListener('click', function(e) {
    if (e.target === this) closeDetail();
  });
});

/* ── Load main data ── */
async function loadData() {
  try {
    const res  = await fetch(`/api/bi/summary/${TOKEN}`);
    const json = await res.json();
    if (!res.ok) throw new Error(json.error || '加载失败');

    allData = json.data || [];

    // Find max pct for bar scaling
    maxPct = Math.max(...allData.map(r => +(r['全局成本占比%'] || 0)), 0.001);

    // KPI
    const totalCost = allData.reduce((s, r) => s + (r['全局总成本贡献'] || 0), 0);
    const maxProd   = Math.max(...allData.map(r => +(r['涉及产品数'] || 0)), 0);
    document.getElementById('kpi-total').textContent    = fmt(allData.length);
    document.getElementById('kpi-max-prod').textContent = fmt(maxProd);
    document.getElementById('kpi-total-cost').textContent = fmtCost(totalCost);
    document.getElementById('nav-total').textContent    = fmt(allData.length);

    // Estimate distinct products
    const maxProdVal = Math.max(...allData.map(r => +(r['涉及产品数'] || 0)), 0);
    document.getElementById('nav-products').textContent = maxProdVal > 0 ? `≥${maxProdVal}` : '—';

    // Range hint: 平均用量占比%
    const avgPctVals = allData.map(r => +(r['平均用量占比%（在各产品中）'] || 0)).filter(v => v > 0);
    if (avgPctVals.length) {
      const minV = Math.min(...avgPctVals).toFixed(1);
      const maxV = Math.max(...avgPctVals).toFixed(1);
      document.getElementById('pct-range-hint').textContent = `(${minV}%–${maxV}%)`;
    }
    // Range hint: 全局总成本贡献
    const costVals = allData.map(r => +(r['全局总成本贡献'] || 0)).filter(v => v > 0);
    if (costVals.length) {
      const minC = Math.min(...costVals);
      const maxC = Math.max(...costVals);
      document.getElementById('cost-range-hint').textContent =
        `(${fmtCost(minC)}–${fmtCost(maxC)})`;
    }
    // Range hint: 全局成本占比%
    const costPctVals = allData.map(r => +(r['全局成本占比%'] || 0)).filter(v => v > 0);
    if (costPctVals.length) {
      const minC = Math.min(...costPctVals);
      const maxC = Math.max(...costPctVals);
      document.getElementById('cost-pct-range-hint').textContent =
        `(${fmtPct(minC)}–${fmtPct(maxC)})`;
    }

    show('loading-state', false);
    show('main-content', true);
    applyFilters();

    // 检查 Excel 是否已就绪，若还在写入则显示"准备中"并轮询
    pollExcelReady();

  } catch (err) {
    show('loading-state', false);
    document.getElementById('error-msg').textContent = String(err);
    show('error-state', true);
  }
}

/* ── Excel 就绪轮询 ── */
async function pollExcelReady() {
  const btn     = document.getElementById('export-btn');
  const icon    = document.getElementById('export-icon');
  const spinner = document.getElementById('export-spinner');
  const label   = document.getElementById('export-label');

  async function check() {
    try {
      const bi = await fetch(`/api/bi_ready/${TOKEN}`);
      if (bi.ok && (await bi.json()).ready === true) {
        const ex = await fetch(`/api/excel_ready/${TOKEN}`);
        if (ex.ok) return (await ex.json()).ready === true;
        return true;
      }
      return false;
    } catch { return false; }
  }

  if (await check()) return;   // 已就绪，无需等待

  // 进入"准备中"状态
  btn.removeAttribute('href');
  btn.style.opacity = '0.7';
  btn.style.cursor  = 'not-allowed';
  icon.classList.add('hidden');
  spinner.classList.remove('hidden');
  label.textContent = '准备中…';

  // 每 2 秒轮询一次，最多等 3 分钟
  let waited = 0;
  const iv = setInterval(async () => {
    waited += 2;
    if (await check() || waited > 180) {
      clearInterval(iv);
      // 恢复按钮
      btn.href          = `/api/download/${TOKEN}`;
      btn.style.opacity = '';
      btn.style.cursor  = '';
      icon.classList.remove('hidden');
      spinner.classList.add('hidden');
      label.textContent = '导出 Excel';
    }
  }, 2000);
}

/* ── Filters ── */
function applyFilters() {
  const search       = (document.getElementById('search-input').value || '').trim().toLowerCase();
  const minProd      = parseFloat(document.getElementById('min-prod-filter').value) || 0;
  const minAvgPct    = parseFloat(document.getElementById('pct-filter').value) || 0;
  const minCost      = parseFloat(document.getElementById('cost-filter').value) || 0;
  const minCostPct   = parseFloat(document.getElementById('cost-pct-filter').value) || 0;

  filteredData = allData.filter(row => {
    if (search     && !String(row['组件编码'] || '').toLowerCase().includes(search)) return false;
    if (minProd    > 0 && (+(row['涉及产品数'] || 0)) < minProd) return false;
    if (minAvgPct  > 0 && (+(row['平均用量占比%（在各产品中）'] || 0)) < minAvgPct) return false;
    if (minCost    > 0 && (+(row['全局总成本贡献'] || 0)) < minCost) return false;
    if (minCostPct > 0 && (+(row['全局成本占比%'] || 0)) < minCostPct) return false;
    return true;
  });

  const isFiltered = search || minProd > 0 || minAvgPct > 0 || minCost > 0 || minCostPct > 0;
  const hint = document.getElementById('filter-hint');
  if (isFiltered) {
    hint.textContent = `已筛选：显示 ${filteredData.length} / ${allData.length} 个组件`;
    hint.classList.remove('hidden');
  } else {
    hint.classList.add('hidden');
  }

  sortAndRender();
}

function resetFilters() {
  document.getElementById('search-input').value    = '';
  document.getElementById('min-prod-filter').value = '';
  document.getElementById('pct-filter').value      = '';
  document.getElementById('cost-filter').value     = '';
  document.getElementById('cost-pct-filter').value = '';
  applyFilters();
}

/* ── Sort ── */
function sortBy(col) {
  if (sortCol === col) {
    sortAsc = !sortAsc;
  } else {
    sortCol = col;
    sortAsc = false;
  }
  sortAndRender();
}

function sortAndRender() {
  const asc = sortAsc;
  const col = sortCol;
  filteredData.sort((a, b) => {
    const av = a[col], bv = b[col];
    if (av === null || av === undefined) return 1;
    if (bv === null || bv === undefined) return -1;
    if (typeof av === 'number' && typeof bv === 'number') return asc ? av - bv : bv - av;
    return asc
      ? String(av).localeCompare(String(bv), 'zh')
      : String(bv).localeCompare(String(av), 'zh');
  });

  // Update sort icons
  ['涉及产品数', '全局总成本贡献', '全局总用量', '平均用量占比%（在各产品中）', '全局成本占比%'].forEach(c => {
    const el = document.getElementById('sort-icon-' + c);
    if (!el) return;
    if (c === col) {
      el.textContent = asc ? '↑' : '↓';
      el.className = 'ml-1 text-indigo-500';
    } else {
      el.textContent = '↕';
      el.className = 'ml-1 opacity-30';
    }
  });

  renderTable();
}

/* ── Render main table ── */
function renderTable() {
  const tbody = document.getElementById('main-tbody');
  const empty = document.getElementById('table-empty');
  const footer = document.getElementById('table-footer');
  const footerCount = document.getElementById('footer-count');
  const kpiShown = document.getElementById('kpi-shown');

  kpiShown.textContent = fmt(filteredData.length);

  if (filteredData.length === 0) {
    tbody.innerHTML = '';
    show('table-empty', true);
    footerCount.textContent = '共 0 行';
    return;
  }
  show('table-empty', false);
  footerCount.textContent = `共 ${fmt(filteredData.length)} 行`;

  tbody.innerHTML = filteredData.map((row, i) => {
    const comp     = esc(row['组件编码'] || '—');
    const prodCnt  = +(row['涉及产品数'] || 0);
    const cost     = +(row['全局总成本贡献'] || 0);
    const qty      = +(row['全局总用量'] || 0);
    const avgPct   = +(row['平均用量占比%（在各产品中）'] || 0);
    const pct      = +(row['全局成本占比%'] || 0);
    // 绝对比例：pct 本身已是 0-100 的百分比数值，直接作为进度条宽度
    // 极小非零值保证最小可见宽度 1.5%，让用户知道"有值"
    const barW     = pct > 0 ? Math.max(1.5, Math.min(100, pct)) : 0;
    const bg       = i % 2 === 0 ? 'bg-white' : 'bg-slate-50/50';

    return `<tr class="trow ${bg}">
      <td class="px-4 py-3 nowrap font-mono text-xs font-semibold text-slate-700">${comp}</td>
      <td class="px-4 py-3 text-center">
        <button class="count-badge" data-comp="${esc(String(row['组件编码'] ?? ''))}">
          ${fmt(prodCnt)}
        </button>
      </td>
      <td class="px-4 py-3 nowrap text-right text-slate-700">${fmtCost(cost)}</td>
      <td class="px-4 py-3 nowrap text-right text-slate-600">${fmtNum(qty)}</td>
      <td class="px-4 py-3 nowrap text-right text-slate-600">${fmtPct(avgPct)}</td>
      <td class="px-4 py-3 nowrap min-w-[200px]">
        <div class="flex items-center gap-2">
          <span class="text-xs text-slate-700 w-16 text-right flex-shrink-0">${fmtPct(pct)}</span>
          <div class="pct-bar-bg flex-1" title="全局成本占比：${pct}%">
            <div class="pct-bar-fill" style="width:${barW.toFixed(2)}%"></div>
          </div>
        </div>
      </td>
    </tr>`;
  }).join('');
}

/* ── Show detail modal ── */
async function showDetail(componentCode) {
  const comp = String(componentCode);
  document.getElementById('modal-component-code').textContent = comp;
  document.getElementById('modal-component-name').textContent = '';
  document.getElementById('modal-count-info').textContent = '—';
  document.getElementById('detail-pct-filter').value = '';
  document.getElementById('detail-tbody').innerHTML = '';
  document.getElementById('detail-footer-count').textContent = '加载中…';
  show('detail-empty', false);
  show('detail-loading', true);
  show('detail-modal', true);
  document.body.style.overflow = 'hidden';

  try {
    const url = `/api/bi/detail/${TOKEN}?component=${encodeURIComponent(comp)}`;
    const res  = await fetch(url);
    const json = await res.json();
    if (!res.ok) throw new Error(json.error || '加载失败');

    detailRaw = json.data || [];
    show('detail-loading', false);

    // Check if name/unit columns exist
    hasName = detailRaw.some(r => r['组件名称'] !== null && r['组件名称'] !== undefined && r['组件名称'] !== '');
    hasUnit = detailRaw.some(r => r['计量单位'] !== null && r['计量单位'] !== undefined && r['计量单位'] !== '');

    // Show component name in header if available
    if (hasName && detailRaw.length > 0) {
      const name = detailRaw[0]['组件名称'];
      if (name) document.getElementById('modal-component-name').textContent = name;
    }

    toggleEl('dth-name', hasName);
    toggleEl('dth-unit', hasUnit);
    document.getElementById('modal-count-info').textContent = json.count;

    applyDetailFilter();
  } catch (err) {
    show('detail-loading', false);
    toast('加载详情失败：' + String(err), 'error');
  }
}

/* ── Detail filter ── */
function applyDetailFilter() {
  const minPct = parseFloat(document.getElementById('detail-pct-filter').value) || 0;
  detailFiltered = detailRaw.filter(row => {
    if (minPct > 0 && (+(row['用量占比%'] || 0)) < minPct) return false;
    return true;
  });
  renderDetailTable();
}

function resetDetailFilter() {
  document.getElementById('detail-pct-filter').value = '';
  applyDetailFilter();
}

function renderDetailTable() {
  const tbody = document.getElementById('detail-tbody');
  const empty = document.getElementById('detail-empty');
  const footerCount = document.getElementById('detail-footer-count');

  if (detailFiltered.length === 0) {
    tbody.innerHTML = '';
    show('detail-empty', true);
    footerCount.textContent = '0 条记录';
    return;
  }
  show('detail-empty', false);
  footerCount.textContent = `共 ${fmt(detailFiltered.length)} 条 / 原始 ${fmt(detailRaw.length)} 条`;

  tbody.innerHTML = detailFiltered.map((row, i) => {
    const bg   = i % 2 === 0 ? 'bg-white' : 'bg-slate-50/50';
    const prod = esc(row['产品编码'] || '—');
    const name = esc(row['组件名称'] || '—');
    const unit = esc(row['计量单位'] || '—');
    const qty  = +(row['MENGE合计'] || 0);
    const pct  = +(row['用量占比%'] || 0);
    const cost = +(row['行原材料成本'] || 0);
    const up   = +(row['组件单价'] || 0);

    return `<tr class="trow ${bg}">
      <td class="px-4 py-2.5 font-mono text-xs font-semibold text-slate-700 nowrap">${prod}</td>
      ${hasName ? `<td class="px-4 py-2.5 text-xs text-slate-600 nowrap">${name}</td>` : ''}
      ${hasUnit ? `<td class="px-4 py-2.5 text-xs text-slate-500 nowrap">${unit}</td>` : ''}
      <td class="px-4 py-2.5 text-right nowrap font-medium text-slate-800">${fmtNum(qty)}</td>
      <td class="px-4 py-2.5 text-right nowrap">
        <span class="inline-block bg-indigo-50 text-indigo-700 text-xs px-2 py-0.5 rounded-full font-medium">${fmtPct(pct)}</span>
      </td>
      <td class="px-4 py-2.5 text-right nowrap text-slate-700">${fmtCost(cost)}</td>
      <td class="px-4 py-2.5 text-right nowrap text-slate-500 text-xs">${fmtNum(up)}</td>
    </tr>`;
  }).join('');
}

function closeDetail() {
  show('detail-modal', false);
  document.body.style.overflow = '';
  detailRaw = [];
  detailFiltered = [];
}

/* ── Helpers ── */
function show(id, visible) {
  const el = document.getElementById(id);
  if (el) el.classList.toggle('hidden', !visible);
}

function toggleEl(id, visible) {
  const el = document.getElementById(id);
  if (el) el.classList.toggle('hidden', !visible);
}

function fmt(n) {
  return Number(n).toLocaleString('zh-CN');
}

function fmtNum(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const num = +n;
  if (Math.abs(num) >= 1e6) return (num / 1e6).toFixed(2) + ' M';
  if (Math.abs(num) >= 1e3) return num.toLocaleString('zh-CN', {maximumFractionDigits: 2});
  return num.toLocaleString('zh-CN', {maximumFractionDigits: 4});
}

/* 自适应精度百分比：避免将 0.0063% 显示为 0.00% */
function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const v = +n;
  if (v === 0) return '0%';
  if (v >= 0.01)   return v.toFixed(2) + '%';
  if (v >= 0.001)  return v.toFixed(3) + '%';
  if (v >= 0.0001) return v.toFixed(4) + '%';
  return '<0.0001%';
}

function fmtCost(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const num = +n;
  if (Math.abs(num) >= 1e8) return (num / 1e8).toFixed(2) + ' 亿';
  if (Math.abs(num) >= 1e4) return (num / 1e4).toFixed(2) + ' 万';
  return num.toLocaleString('zh-CN', {maximumFractionDigits: 2});
}

function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}

/* ── Toast ── */
function toast(msg, type = 'info') {
  const wrap  = document.getElementById('toast-wrap');
  const el    = document.createElement('div');
  const color = { success: 'bg-green-600', error: 'bg-red-600', info: 'bg-slate-700' }[type] || 'bg-slate-700';
  el.className = `t-in pointer-events-auto flex items-center gap-3 ${color} text-white text-sm font-medium px-4 py-3 rounded-xl shadow-lg max-w-xs`;
  el.innerHTML = `<span>${esc(msg)}</span>`;
  wrap.appendChild(el);
  setTimeout(() => {
    el.classList.replace('t-in', 't-out');
    setTimeout(() => el.remove(), 300);
  }, 4000);
}

/* ══════════════════════════════════════════════════════════
   Tab 切换
══════════════════════════════════════════════════════════ */
function switchTab(name) {
  ['cost', 'price'].forEach(t => {
    document.getElementById('tab-' + t + '-content').classList.toggle('hidden', t !== name);
    const btn = document.getElementById('tab-btn-' + t);
    if (btn) btn.classList.toggle('active', t === name);
  });
}

/* ══════════════════════════════════════════════════════════
   价格波动分析
══════════════════════════════════════════════════════════ */
let pfData         = [];
let pfCatData      = [];   // raw category impact from API
let pfHasCategory  = false;
let pfHasName      = false;
let pfSortCol      = '价格差值';
let pfSortAsc      = false;   // false = 降序（默认：最大涨幅在前）

async function queryPriceFluctuation() {
  const rawStart = document.getElementById('pf-start-date').value;
  const rawEnd   = document.getElementById('pf-end-date').value;
  const minPct   = document.getElementById('pf-min-pct').value || '0';

  if (!rawStart || !rawEnd) {
    toast('请先选择开始日期和结束日期', 'error');
    return;
  }
  if (rawStart > rawEnd) {
    toast('开始日期不能晚于结束日期', 'error');
    return;
  }

  const startDate = rawStart.replace(/-/g, '');
  const endDate   = rawEnd.replace(/-/g, '');

  show('pf-init-hint', false);
  show('pf-loading', true);
  show('pf-results', false);

  try {
    const url = `/api/bi/price_fluctuation/${TOKEN}?start_date=${startDate}&end_date=${endDate}&min_change_pct=${minPct}`;
    const res  = await fetch(url);
    const json = await res.json();

    show('pf-loading', false);

    if (!res.ok) {
      toast((json.error || '查询失败'), 'error');
      show('pf-init-hint', true);
      return;
    }

    if (json.message) {
      toast(json.message, 'info');
    }

    pfData        = json.data          || [];
    pfHasCategory = json.has_category  || false;
    pfHasName     = json.has_name      || false;
    pfSortCol     = '价格差值';
    pfSortAsc     = false;
    // 重置列头图标
    ['价格差值', '价格变动%'].forEach(c => {
      const el = document.getElementById('pf-sort-icon-' + c);
      if (!el) return;
      el.textContent = c === '价格差值' ? '↓' : '↕';
      el.className   = c === '价格差值' ? 'ml-0.5 text-indigo-500' : 'ml-0.5 opacity-30';
    });

    renderPfKpi(json);
    renderPfTable();
    renderPfCategoryImpact(json.category_impact || []);
    renderPfAdvice(json.category_impact || []);

    show('pf-results', true);

  } catch (err) {
    show('pf-loading', false);
    show('pf-init-hint', true);
    toast('查询失败：' + String(err), 'error');
  }
}

function resetPriceFilter() {
  document.getElementById('pf-start-date').value = '';
  document.getElementById('pf-end-date').value   = '';
  document.getElementById('pf-min-pct').value    = '';
  show('pf-results', false);
  show('pf-init-hint', true);
  pfData    = [];
  pfCatData = [];
  pfSortCol = '价格差值';
  pfSortAsc = false;
}

async function queryAiPurchaseAdvice() {
  const ai = (CAPABILITIES && CAPABILITIES.ai_purchase_advice) || {};
  if (!ai.enabled) {
    toast(ai.reason || 'AI 采购建议未配置', 'error');
    return;
  }

  // 防重复点击：按钮处于 loading 状态时忽略
  const btn = document.getElementById('ai-advice-btn');
  if (btn && (btn.disabled || btn.dataset.loading === '1')) return;
  if (btn) { btn.dataset.loading = '1'; btn.style.opacity = '0.6'; btn.style.cursor = 'not-allowed'; }

  show('ai-advice-panel', false);
  show('ai-advice-loading', true);
  const adviceTextEl = document.getElementById('ai-advice-text');
  adviceTextEl.innerHTML = '';

  function resetBtn() {
    if (btn) { btn.dataset.loading = ''; btn.style.opacity = ''; btn.style.cursor = ''; }
  }

  try {
    const res = await fetch(`/api/bi/ai_purchase_advice_stream/${TOKEN}`);
    if (!res.ok) {
      const json = await res.json().catch(() => ({}));
      if (json.code === 'AI_DISABLED') applyAiCapabilities();
      toast(json.error || 'AI 建议生成失败', 'error');
      show('ai-advice-loading', false);
      resetBtn();
      return;
    }

    show('ai-advice-loading', false);
    show('ai-advice-panel', true);

    // 等待 DeepSeek 开始输出前的占位提示（通常需 15–30s 工具调用）
    adviceTextEl.innerHTML =
      '<p class="text-amber-500 text-sm animate-pulse">⏳ DeepSeek 正在联网查询行情，请稍候（通常需要 15–30 秒）…</p>';

    const reader = res.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let buffer = '';
    let markdown = '';
    let hadError = false;
    let streamPre = null;   // 流式阶段用 <pre> 追加纯文本，避免 innerHTML 闪烁

    function handleSseBlock(block) {
      if (!block.trim()) return;
      let event = 'message';
      const dataLines = [];
      for (const line of block.split(/\r?\n/)) {
        if (line.startsWith('event:')) event = line.slice(6).trim();
        if (line.startsWith('data:')) dataLines.push(line.slice(5).trim());
      }
      if (!dataLines.length) return;
      const data = JSON.parse(dataLines.join('\n'));

      if (event === 'meta') {
        const range = data.date_range || {};
        document.getElementById('ai-advice-range').textContent =
          range.start && range.end ? `分析区间：${range.start} 至 ${range.end}` : '分析区间：最近三个月';
        document.getElementById('ai-advice-total').textContent = `共 ${fmt(data.total || 0)} 个物料`;
        renderAiMaterialsPreview(data.materials || []);

      } else if (event === 'delta') {
        // 第一个 delta 到来：清除占位符，创建 <pre> 用于无闪烁追加
        if (!streamPre) {
          adviceTextEl.innerHTML = '';
          streamPre = document.createElement('pre');
          streamPre.className = 'whitespace-pre-wrap text-sm text-slate-700 leading-7 font-sans';
          adviceTextEl.appendChild(streamPre);
        }
        const chunk = data.text || '';
        markdown += chunk;
        // textContent 追加不触发 HTML 解析，零闪烁
        streamPre.textContent += chunk;
        // 自动滚到最新文字
        adviceTextEl.scrollTop = adviceTextEl.scrollHeight;

      } else if (event === 'error') {
        hadError = true;
        if (data.code === 'AI_DISABLED') applyAiCapabilities();
        toast(data.error || 'AI 建议生成失败', 'error');

      } else if (event === 'done') {
        // 流结束：用完整 Markdown 渲染替换 <pre> 纯文本
        adviceTextEl.innerHTML = renderMarkdown(markdown);
        if (!hadError) toast('AI 采购建议已生成', 'success');
        resetBtn();
      }
    }

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const parts = buffer.split(/\r?\n\r?\n/);
      buffer = parts.pop() || '';
      for (const part of parts) handleSseBlock(part);
    }
    // 处理最后一块不完整的 SSE（流意外断开时）
    if (buffer.trim()) handleSseBlock(buffer);
    // 流异常结束时仍尝试渲染已有内容
    if (markdown && !adviceTextEl.querySelector(':not(pre)')) {
      adviceTextEl.innerHTML = renderMarkdown(markdown);
    }
    resetBtn();

  } catch (err) {
    show('ai-advice-loading', false);
    toast('AI 建议生成失败：' + String(err), 'error');
    resetBtn();
  }
}

function renderAiMaterialsPreview(materials) {
  const rows = materials.slice(0, 10);
  const tbody = document.getElementById('ai-materials-tbody');
  if (!rows.length) {
    tbody.innerHTML = '';
    show('ai-materials-preview', false);
    return;
  }

  tbody.innerHTML = rows.map(r => {
    return `<tr>
      <td class="px-3 py-2 text-slate-700 nowrap">${esc(r['物料名称'] || '—')}</td>
      <td class="px-3 py-2 font-mono text-slate-600 nowrap">${esc(r['组件编码'] || '—')}</td>
      <td class="px-3 py-2 text-right font-mono text-slate-500 nowrap">${fmt(r['涉及产品数'] || 0)}</td>
      <td class="px-3 py-2 text-right font-mono text-slate-500 nowrap">${fmt(r['记录数'] || 0)}</td>
    </tr>`;
  }).join('');
  show('ai-materials-preview', true);
}

function renderMarkdown(md) {
  const lines = esc(md).split(/\r?\n/);
  let html = '';
  let inList = false;
  let inTable = false;

  function inline(text) {
    return text
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  }

  function closeTable() {
    if (inTable) {
      html += '</tbody></table></div>';
      inTable = false;
    }
  }

  function closeList() {
    if (inList) {
      html += '</ul>';
      inList = false;
    }
  }

  function isTableSeparator(line) {
    return /^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$/.test(line);
  }

  function parseTableCells(line) {
    return line
      .replace(/^\|/, '')
      .replace(/\|$/, '')
      .split('|')
      .map(cell => inline(cell.trim()));
  }

  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    const line = raw.trim();
    if (!line) {
      closeList();
      closeTable();
      html += '<br>';
      continue;
    }
    if (/^-{3,}$/.test(line)) continue;

    const nextLine = (lines[i + 1] || '').trim();
    if (line.includes('|') && nextLine && isTableSeparator(nextLine)) {
      closeList();
      closeTable();
      const headers = parseTableCells(line);
      html += '<div class="md-table-wrap"><table><thead><tr>';
      html += headers.map(h => `<th>${h}</th>`).join('');
      html += '</tr></thead><tbody>';
      inTable = true;
      i++; // skip separator
      continue;
    }

    if (inTable && line.includes('|') && !isTableSeparator(line)) {
      const cells = parseTableCells(line);
      html += '<tr>' + cells.map(c => `<td>${c}</td>`).join('') + '</tr>';
      continue;
    }

    closeTable();

    if (line.startsWith('## ')) {
      closeList();
      html += `<h2>${inline(line.slice(3))}</h2>`;
    } else if (line.startsWith('### ')) {
      closeList();
      html += `<h3>${inline(line.slice(4))}</h3>`;
    } else if (/^[-*]\s+/.test(line) || /^\d+\.\s+/.test(line)) {
      if (!inList) { html += '<ul>'; inList = true; }
      html += `<li>${inline(line.replace(/^[-*]\s+/, '').replace(/^\d+\.\s+/, ''))}</li>`;
    } else {
      closeList();
      html += `<p>${inline(line)}</p>`;
    }
  }
  closeList();
  closeTable();
  return html;
}

function pfSortBy(col) {
  if (pfSortCol === col) {
    pfSortAsc = !pfSortAsc;
  } else {
    pfSortCol = col;
    pfSortAsc = false;
  }
  // 更新列头图标
  ['价格差值', '价格变动%'].forEach(c => {
    const el = document.getElementById('pf-sort-icon-' + c);
    if (!el) return;
    if (c === pfSortCol) {
      el.textContent = pfSortAsc ? '↑' : '↓';
      el.className = 'ml-0.5 text-indigo-500';
    } else {
      el.textContent = '↕';
      el.className = 'ml-0.5 opacity-30';
    }
  });
  renderPfTable();
}

function renderPfKpi(json) {
  document.getElementById('pf-kpi-total').textContent = fmt(json.total || 0);

  const dr = json.date_range || {};
  const fmt8 = d => d ? `${d.slice(0,4)}-${d.slice(4,6)}-${d.slice(6,8)}` : '—';
  document.getElementById('pf-kpi-range').textContent =
    (dr.start && dr.end) ? `${fmt8(dr.start)} → ${fmt8(dr.end)}` : '—';

  const ups   = pfData.filter(r => r['价格变动%'] > 0).sort((a,b) => b['价格变动%'] - a['价格变动%']);
  const downs = pfData.filter(r => r['价格变动%'] < 0).sort((a,b) => a['价格变动%'] - b['价格变动%']);

  if (ups.length > 0) {
    document.getElementById('pf-kpi-max-up').textContent =
      `${ups[0]['组件编码']} (+${ups[0]['价格变动%'].toFixed(1)}%)`;
  } else {
    document.getElementById('pf-kpi-max-up').textContent = '无涨价';
    document.getElementById('pf-kpi-max-up').className = 'text-lg font-bold text-slate-400 truncate';
  }

  if (downs.length > 0) {
    document.getElementById('pf-kpi-max-down').textContent =
      `${downs[0]['组件编码']} (${downs[0]['价格变动%'].toFixed(1)}%)`;
  } else {
    document.getElementById('pf-kpi-max-down').textContent = '无降价';
    document.getElementById('pf-kpi-max-down').className = 'text-lg font-bold text-slate-400 truncate';
  }
}

function renderPfTable() {
  const tbody      = document.getElementById('pf-tbody');
  const empty      = document.getElementById('pf-table-empty');
  const footerCount = document.getElementById('pf-footer-count');

  toggleEl('pf-th-name', pfHasName);
  toggleEl('pf-th-cat',  pfHasCategory);

  // 按当前排序列排序
  const col = pfSortCol;
  const asc = pfSortAsc;
  pfData.sort((a, b) => {
    const av = a[col] ?? -Infinity;
    const bv = b[col] ?? -Infinity;
    return asc ? av - bv : bv - av;
  });

  if (pfData.length === 0) {
    tbody.innerHTML = '';
    show('pf-table-empty', true);
    footerCount.textContent = '共 0 行';
    return;
  }
  show('pf-table-empty', false);
  footerCount.textContent = `共 ${fmt(pfData.length)} 行`;

  tbody.innerHTML = pfData.map((row, i) => {
    const pct       = +(row['价格变动%'] || 0);
    const diff      = +(row['价格差值']  || 0);
    const bg        = i % 2 === 0 ? 'bg-white' : 'bg-slate-50/50';
    const pctClass  = pct > 0 ? 'badge-up' : (pct < 0 ? 'badge-down' : 'badge-flat');
    const pctArrow  = pct > 0 ? '▲' : (pct < 0 ? '▼' : '—');
    const diffColor = diff > 0 ? 'text-red-600' : (diff < 0 ? 'text-green-600' : 'text-slate-500');
    const diffSign  = diff > 0 ? '+' : '';

    return `<tr class="trow ${bg}" data-comp="${esc(String(row['组件编码'] || ''))}">
      <td class="px-4 py-2.5 text-xs text-slate-400">${i + 1}</td>
      <td class="px-4 py-2.5 font-mono text-xs font-semibold text-slate-700 nowrap">${esc(row['组件编码'] || '—')}</td>
      ${pfHasName ? `<td class="px-4 py-2.5 text-xs text-slate-600 nowrap">${esc(row['组件名称'] || '—')}</td>` : ''}
      ${pfHasCategory ? `<td class="px-4 py-2.5 text-xs nowrap">
        <span class="inline-block px-2 py-0.5 bg-indigo-50 text-indigo-700 rounded-full text-xs">${esc(row['分类'] || '—')}</span>
      </td>` : ''}
      <td class="px-4 py-2.5 text-right nowrap text-slate-700 text-xs font-mono">${fmtPrice(row['期初价格'])}</td>
      <td class="px-4 py-2.5 text-right nowrap text-slate-700 text-xs font-mono">${fmtPrice(row['期末价格'])}</td>
      <td class="px-4 py-2.5 text-right nowrap text-xs font-mono font-semibold ${diffColor}">${diffSign}${fmtPrice(diff)}</td>
      <td class="px-4 py-2.5 text-right nowrap">
        <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold ${pctClass}">
          ${pctArrow} ${Math.abs(pct).toFixed(2)}%
        </span>
      </td>
      <td class="px-4 py-2.5 text-right nowrap text-slate-500 text-xs font-mono">${fmtPrice(row['期间最低价'])}</td>
      <td class="px-4 py-2.5 text-right nowrap text-slate-500 text-xs font-mono">${fmtPrice(row['期间最高价'])}</td>
      <td class="px-4 py-2.5 text-center text-xs text-slate-600">${row['涉及产品数'] || 0}</td>
    </tr>`;
  }).join('');
}

/* 根据阈值计算影响等级（客户端，支持用户调整） */
function computeLevel(absMax) {
  const high = parseFloat(document.getElementById('pf-level-high').value) || 10;
  const mid  = parseFloat(document.getElementById('pf-level-mid').value)  || 3;
  return absMax >= high ? '高' : (absMax >= mid ? '中' : '低');
}

function renderPfCategoryImpact(catImpact) {
  pfCatData = catImpact;   // 保存原始数据供阈值变更时重新渲染
  _drawCatTable();
}

function recomputeImpactLevels() {
  if (!pfCatData.length) return;
  _drawCatTable();
  // 同步更新建议区（影响等级可能变化）
  renderPfAdvice(pfCatData);
}

function _drawCatTable() {
  const section = document.getElementById('pf-cat-section');
  const tbody   = document.getElementById('pf-cat-tbody');

  if (!pfHasCategory || pfCatData.length === 0) {
    show('pf-cat-section', false);
    return;
  }

  const levelClass = { '高': 'level-high', '中': 'level-mid', '低': 'level-low' };

  tbody.innerHTML = pfCatData.map((row, i) => {
    const bg     = i % 2 === 0 ? 'bg-white' : 'bg-slate-50/50';
    const avg    = +(row['平均变动%']    || 0);
    const maxUp  = +(row['最大涨幅%']    || 0);
    const maxDn  = +(row['最大跌幅%']    || 0);
    const level  = computeLevel(Math.abs(avg));   // 用平均变动%绝对值与阈值比较
    const lc     = levelClass[level] || 'level-low';
    const catVal = esc(row['分类']);

    return `<tr class="trow ${bg}">
      <td class="px-4 py-2.5 font-semibold text-slate-700 text-sm nowrap">${catVal}</td>
      <td class="px-4 py-2.5 text-right nowrap text-slate-600">${fmt(row['涉及组件数'])}</td>
      <td class="px-4 py-2.5 text-right nowrap font-medium ${avg >= 0 ? 'text-red-600' : 'text-green-600'}">${avg >= 0 ? '+' : ''}${avg.toFixed(2)}%</td>
      <td class="px-4 py-2.5 text-right nowrap text-red-600 font-medium jump-cell"
          onclick="jumpToComp('${catVal}','up')" title="点击定位该分类最大涨幅组件">
        +${maxUp.toFixed(2)}%
      </td>
      <td class="px-4 py-2.5 text-right nowrap text-green-600 font-medium jump-cell"
          onclick="jumpToComp('${catVal}','down')" title="点击定位该分类最大跌幅组件">
        ${maxDn.toFixed(2)}%
      </td>
      <td class="px-4 py-2.5 text-right nowrap text-slate-500 text-xs">${fmt(row['上涨组件数'])} ↑ / ${fmt(row['下跌组件数'])} ↓</td>
      <td class="px-4 py-2.5 text-center nowrap">
        <span class="inline-block px-2 py-0.5 rounded-full text-xs font-semibold ${lc}">${level}</span>
      </td>
    </tr>`;
  }).join('');

  show('pf-cat-section', true);
}

/* 跳转到分类中最大涨幅（up）或最大跌幅（down）的组件行 */
function jumpToComp(cat, direction) {
  const candidates = pfData.filter(r => String(r['分类'] || '') === cat);
  if (!candidates.length) { toast('该分类下暂无组件数据', 'info'); return; }

  const target = direction === 'up'
    ? candidates.reduce((best, r) => (+(r['价格变动%'] || 0)) > (+(best['价格变动%'] || 0)) ? r : best)
    : candidates.reduce((best, r) => (+(r['价格变动%'] || 0)) < (+(best['价格变动%'] || 0)) ? r : best);

  const compCode = String(target['组件编码'] || '');
  const rows = document.querySelectorAll('#pf-tbody tr[data-comp]');
  let targetRow = null;
  for (const row of rows) {
    if (row.dataset.comp === compCode) { targetRow = row; break; }
  }

  if (!targetRow) {
    toast(`未找到组件 ${compCode}（可能已被百分比过滤隐藏）`, 'info');
    return;
  }

  // 清除上一次高亮（避免重复点击时动画残留）
  targetRow.classList.remove('pf-highlight');
  void targetRow.offsetWidth;

  // 用 IntersectionObserver 监听行真正进入视口后再触发高亮，
  // 确保滚动动画完成时高亮才开始，不受滚动距离影响。
  const observer = new IntersectionObserver((entries, obs) => {
    const entry = entries[0];
    if (entry.isIntersecting) {
      obs.disconnect();
      // 再次强制 reflow，防止 remove+add 在同一帧被合并
      void entry.target.offsetWidth;
      entry.target.classList.add('pf-highlight');
      setTimeout(() => entry.target.classList.remove('pf-highlight'), 3600);
    }
  }, { threshold: 0.6 });   // 行至少 60% 可见时才触发

  observer.observe(targetRow);
  targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });

  // 安全超时：若 5s 内元素始终未进入视口（被遮挡等），取消观察
  setTimeout(() => observer.disconnect(), 5000);
}

function renderPfAdvice(catImpact) {
  const adviceEl = document.getElementById('pf-advice');
  const textEl   = document.getElementById('pf-advice-text');

  if (pfData.length === 0) {
    show('pf-advice', false);
    return;
  }

  const lines = [];

  // 价格上涨的高影响分类（用客户端实时计算的等级）
  if (pfHasCategory) {
    const withLevel = catImpact.map(c => ({
      ...c,
      _level: computeLevel(Math.abs(+(c['平均变动%'] || 0)))
    }));
    const highUp   = withLevel.filter(c => c._level === '高' && c['平均变动%'] > 0);
    const highDown = withLevel.filter(c => c._level === '高' && c['平均变动%'] < 0);
    const midAll   = withLevel.filter(c => c._level === '中');
    const highThreshold = parseFloat(document.getElementById('pf-level-high').value) || 10;
    if (highUp.length > 0) {
      lines.push(`<strong>重点关注：</strong>「${highUp.map(c => esc(c['分类'])).join('」「')}」类组件均价大幅上涨（最大绝对变动 ≥${highThreshold}%），建议优先评估该类别采购成本并考虑锁价或寻源。`);
    }
    if (highDown.length > 0) {
      lines.push(`<strong>降价机会：</strong>「${highDown.map(c => esc(c['分类'])).join('」「')}」类组件价格出现明显下降，可考虑重新谈判合同价格或追加采购量。`);
    }
    if (midAll.length > 0) {
      lines.push(`<strong>持续跟踪：</strong>「${midAll.map(c => esc(c['分类'])).join('」「')}」类组件波动幅度中等，建议每月复查报价。`);
    }
  }

  // 无分类时基于组件数据给出建议
  if (!pfHasCategory || lines.length === 0) {
    const upCnt   = pfData.filter(r => r['价格变动%'] > 0).length;
    const downCnt = pfData.filter(r => r['价格变动%'] < 0).length;
    const bigUp   = pfData.filter(r => r['价格变动%'] >= 10).length;
    if (bigUp > 0) {
      lines.push(`共有 <strong>${bigUp}</strong> 个组件价格涨幅≥10%，建议逐项核实供应商报价合理性，并在产品报价中同步调整成本。`);
    }
    if (upCnt > downCnt) {
      lines.push(`当前区间整体呈涨价趋势（${upCnt} 涨 / ${downCnt} 跌），建议加快推进关键物料锁价。`);
    } else if (downCnt > upCnt) {
      lines.push(`当前区间整体呈降价趋势（${downCnt} 跌 / ${upCnt} 涨），可适时重新议价以降低采购成本。`);
    } else if (pfData.length > 0) {
      lines.push('所选区间内价格波动，请结合业务实际评估采购策略。');
    }
  }

  if (lines.length === 0) {
    show('pf-advice', false);
    return;
  }

  textEl.innerHTML = lines.map(l => `<p>${l}</p>`).join('');
  show('pf-advice', true);
}

function fmtPrice(n) {
  if (n === null || n === undefined || isNaN(+n)) return '—';
  return (+n).toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 4 });
}
