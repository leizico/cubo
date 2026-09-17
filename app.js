// app.js — Bi(Cubo) Enterprise OLAP Engine & Matrix Controller
import { initAuth, getSession } from './auth.js';
import { api } from './api-client.js';

// CALENDAR MONTH ORDERING ENGINE
const MONTH_MAP = {
  'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4, 'maio': 5, 'junho': 6,
  'julho': 7, 'agosto': 8, 'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12,
  'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
  'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12,
  'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
  'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
  'feb': 2, 'apr': 4, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
};

function getMonthIndex(val) {
  if (val === null || val === undefined) return 99;
  const str = String(val).trim().toLowerCase();
  return MONTH_MAP[str] || 99;
}

function isMonthField(fieldName) {
  if (!fieldName) return false;
  return /mês|mes|month|periodo|período/i.test(fieldName);
}

/** Normaliza valores de dimensão (null/vazio → marcador estável). */
function normalizeDimValue(v) {
  if (v === null || v === undefined || v === '') return '(vazio)';
  return v;
}

/** Verifica inclusão em Set tolerando number vs string (ex.: restore do localStorage). */
function selectionHas(sel, val) {
  if (!sel) return false;
  return sel.has(val) || sel.has(String(val));
}

/** Lista distinta ordenada de um campo, incluindo '(vazio)' quando existir. */
function getDistinctValues(field) {
  const distinct = Array.from(new Set(state.data.map((r) => normalizeDimValue(r[field]))));
  if (isMonthField(field)) {
    distinct.sort((a, b) => getMonthIndex(a) - getMonthIndex(b));
  } else {
    distinct.sort((a, b) => ('' + a).localeCompare(('' + b), 'pt-BR', { numeric: true }));
  }
  return distinct;
}

function dimValuesMatch(a, b) {
  return String(normalizeDimValue(a)) === String(normalizeDimValue(b));
}

// AGGREGATION FUNCTIONS
const AGGS = {
  sum: { label: 'Soma', fn: (arr) => arr.reduce((a, b) => a + b, 0) },
  count: { label: 'Contagem', fn: (arr) => arr.length },
  avg: { label: 'Média', fn: (arr) => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0 },
  min: { label: 'Mínimo', fn: (arr) => arr.length ? Math.min(...arr) : 0 },
  max: { label: 'Máximo', fn: (arr) => arr.length ? Math.max(...arr) : 0 },
};

// GLOBAL APP STATE
const state = {
  data: [],
  fields: [],
  fieldTypes: {},
  fieldAliases: {}, // field -> custom alias label
  fieldSortDirs: {}, // field -> 'asc' | 'desc'
  workbook: null,
  activeSheet: null,
  pivot: {
    filters: [],
    rows: [],
    cols: [],
    values: [] // [{field, agg}]
  },
  filterSelected: {}, // field -> Set of allowed values
  chart: { x: null, series: null, value: null, agg: 'sum' },
  chartInstance: null,
  lastMatrixAOA: null,
  collapsedNodes: new Set(), // stores collapsed tree row keys
  showSubtotals: true, // Excel / Google Sheets group subtotals
  layoutMode: 'tabular', // 'tabular' (Excel) or 'compact' (Google Sheets)
  currentDatasetId: null,
  currentDatasetName: null,
};

export function getFieldLabel(field) {
  if (!field) return '';
  return state.fieldAliases[field] || field;
}

// DEMO ENTERPRISE DATASET
function generateDemoDataset() {
  const anos = [2026];
  const meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  const categorias = [
    { familia: 'Eletrônicos', itens: ['Notebook Pro 15', 'Monitor Ultrawide 34', 'Teclado Mecânico', 'Mouse Wireless'] },
    { familia: 'Móveis', itens: ['Cadeira Ergonômica', 'Mesa Reta 1.4m', 'Gaveteiro Volante', 'Armário de Aço'] },
    { familia: 'Vestuário', itens: ['Camisa Polo Executiva', 'Jaqueta de Couro', 'Calça Alfaiataria', 'Blazer Slim'] }
  ];
  const segmentos = ['Varejo', 'B2B Corporativo', 'E-Commerce'];
  const regioes = ['Sudeste', 'Sul', 'Nordeste', 'Centro-Oeste'];
  const vendedores = ['Ana Silva', 'Carlos Souza', 'Mariana Costa', 'Pedro Santos'];

  const rows = [];
  let id = 1001;

  meses.forEach((mes) => {
    categorias.forEach((cat) => {
      cat.itens.forEach((item) => {
        segmentos.forEach((seg) => {
          const reg = regioes[Math.floor(Math.random() * regioes.length)];
          const vend = vendedores[Math.floor(Math.random() * vendedores.length)];
          const qtd = Math.floor(Math.random() * 45) + 5;
          const precoUnit = Math.floor(Math.random() * 350) + 50;
          const fat = qtd * precoUnit;
          const margem = Math.round(fat * (Math.random() * 0.35 + 0.15));

          rows.push({
            'Ano': 2026,
            'Mês': mes,
            'Família': cat.familia,
            'Descrição Item': item,
            'Segmento': seg,
            'Região': reg,
            'Vendedor': vend,
            'Faturamento (R$)': fat,
            'Quantidade (unid)': qtd,
            'Margem (R$)': margem
          });
        });
      });
    });
  });

  return rows;
}

// HELPER FORMATTERS & DETECTOR
function fmtNum(n) {
  if (n === null || n === undefined || isNaN(n)) return '';
  const rounded = Math.round(n * 100) / 100;
  return rounded.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
}

function detectType(values) {
  let numCount = 0, dateCount = 0, total = 0;
  for (const v of values) {
    if (v === null || v === undefined || v === '') continue;
    total++;
    if (typeof v === 'number') numCount++;
    else if (v instanceof Date) dateCount++;
  }
  if (total === 0) return 'text';
  if (dateCount / total > 0.6) return 'date';
  if (numCount / total > 0.6) return 'num';
  return 'text';
}

// LOAD DATASET INTO STATE
async function loadDataset(rows, datasetName = 'Dados Importados', { persist = true, datasetId = null } = {}) {
  state.data = rows;
  state.fields = rows.length ? Object.keys(rows[0]) : [];
  state.fieldTypes = {};
  state.fields.forEach((f) => {
    state.fieldTypes[f] = detectType(rows.map((r) => r[f]));
  });

  state._fieldDistinct = {};
  state.fields.forEach((f) => {
    state._fieldDistinct[f] = new Set(getDistinctValues(f).map(String));
  });

  state.pivot = {
    filters: [],
    rows: ['Família', 'Descrição Item'].filter(f => state.fields.includes(f)),
    cols: ['Mês'].filter(f => state.fields.includes(f)),
    values: state.fields.includes('Faturamento (R$)') 
      ? [{ field: 'Faturamento (R$)', agg: 'sum' }] 
      : (state.fields.length ? [{ field: state.fields.find(f => state.fieldTypes[f] === 'num') || state.fields[0], agg: 'sum' }] : [])
  };

  state.chart = {
    x: state.pivot.cols[0] || state.pivot.rows[0] || null,
    series: state.pivot.rows[0] || null,
    value: state.pivot.values[0] ? state.pivot.values[0].field : null,
    agg: 'sum'
  };

  state.filterSelected = {};
  state.collapsedNodes.clear();
  state.currentDatasetId = datasetId;
  state.currentDatasetName = datasetName;

  document.getElementById('datasetInfoTag').textContent = `Base: ${datasetName} (${rows.length} registros)`;
  document.getElementById('fieldsCountBadge').textContent = state.fields.length;

  refreshAll();

  if (persist && getSession() && rows.length) {
    try {
      const data = await api('/datasets', {
        method: 'POST',
        body: { name: datasetName, rows },
      });
      state.currentDatasetId = data.dataset.id;
      document.getElementById('datasetInfoTag').textContent =
        `Base: ${datasetName} (${rows.length} registros) · nuvem`;
      renderCloudDatasets();
    } catch (err) {
      console.warn('Falha ao persistir dataset:', err.message);
    }
  }
}

// FILE READERS
document.getElementById('fileInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const wb = isCsv
        ? XLSX.read(ev.target.result, { type: 'string', cellDates: true })
        : XLSX.read(ev.target.result, { type: 'array', cellDates: true });
      state.workbook = wb;
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
      
      const select = document.getElementById('sheetSelect');
      if (wb.SheetNames.length > 1) {
        select.innerHTML = wb.SheetNames.map(n => `<option value="${n}">${n}</option>`).join('');
        select.classList.remove('hidden');
      } else {
        select.classList.add('hidden');
      }

      loadDataset(rows, file.name);
    } catch (err) {
      alert('Erro ao carregar arquivo: ' + err.message);
    }
  };
  if (isCsv) reader.readAsText(file, 'UTF-8');
  else reader.readAsArrayBuffer(file);
});

document.getElementById('sheetSelect').addEventListener('change', (e) => {
  if (!state.workbook) return;
  const sheetName = e.target.value;
  const ws = state.workbook.Sheets[sheetName];
  if (!ws) return;
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  loadDataset(rows, `${sheetName}`);
});

document.getElementById('btnLoadDemo').addEventListener('click', () => {
  const demoRows = generateDemoDataset();
  loadDataset(demoRows, 'Vendas Empresariais (Demo)');
});

const btnQuickDemoHero = document.getElementById('btnQuickDemoHero');
if (btnQuickDemoHero) {
  btnQuickDemoHero.addEventListener('click', () => {
    const demoRows = generateDemoDataset();
    loadDataset(demoRows, 'Vendas Empresariais (Demo)');
  });
}

// DRAG & DROP LOGIC
function renderFieldPool() {
  const pool = document.getElementById('fieldPool');
  const searchVal = (document.getElementById('fieldSearchInput').value || '').toLowerCase();
  pool.innerHTML = '';

  const used = new Set([
    ...state.pivot.filters, ...state.pivot.rows, ...state.pivot.cols,
    ...state.pivot.values.map(v => v.field)
  ]);

  const filteredFields = state.fields.filter(f => f.toLowerCase().includes(searchVal));

  if (!filteredFields.length) {
    pool.innerHTML = '<div class="empty-pool-hint">Nenhum campo encontrado.</div>';
    return;
  }

  filteredFields.forEach((f) => {
    const chip = document.createElement('div');
    chip.className = 'field-chip';
    chip.draggable = true;
    chip.dataset.field = f;

    const typeClass = state.fieldTypes[f] === 'num' ? 'num' : (state.fieldTypes[f] === 'date' ? 'date' : 'text');
    const typeLabel = state.fieldTypes[f] === 'num' ? '123' : (state.fieldTypes[f] === 'date' ? '📅' : 'ABC');

    const displayLabel = getFieldLabel(f);
    const subText = f !== displayLabel ? ` <span style="font-size:0.65rem; color:var(--color-text-muted);">(${f})</span>` : '';

    chip.innerHTML = `<span>${displayLabel}${subText}</span><span class="type-tag ${typeClass}">${typeLabel}</span>`;

    if (used.has(f)) chip.style.opacity = '0.4';

    chip.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', JSON.stringify({ field: f, from: 'pool' }));
    });

    pool.appendChild(chip);
  });
}

document.getElementById('fieldSearchInput').addEventListener('input', renderFieldPool);

function setupDropzone(el, zoneKey) {
  el.addEventListener('dragover', (e) => {
    e.preventDefault();
    el.classList.add('dragover');
  });
  el.addEventListener('dragleave', () => el.classList.remove('dragover'));
  el.addEventListener('drop', (e) => {
    e.preventDefault();
    el.classList.remove('dragover');
    let payload;
    try {
      payload = JSON.parse(e.dataTransfer.getData('text/plain'));
    } catch (err) { return; }
    if (!payload || !payload.field) return;
    handleDrop(zoneKey, payload.field);
  });
}

['zoneFilters', 'zoneRows', 'zoneCols', 'zoneValues', 'zoneChartX', 'zoneChartSeries', 'zoneChartValue'].forEach((id) => {
  const el = document.getElementById(id);
  if (el) setupDropzone(el, el.dataset.zone);
});
setupDropzone(document.getElementById('fieldPool'), 'pool');

function removeFieldEverywhere(field, exceptZone) {
  if (exceptZone !== 'rows') removeFrom(state.pivot.rows, field);
  if (exceptZone !== 'cols') removeFrom(state.pivot.cols, field);
  if (exceptZone !== 'filters') removeFrom(state.pivot.filters, field);
  if (exceptZone !== 'values') state.pivot.values = state.pivot.values.filter(v => v.field !== field);
}

function removeFrom(arr, field) {
  const i = arr.indexOf(field);
  if (i > -1) arr.splice(i, 1);
}

function handleDrop(zoneKey, field) {
  if (zoneKey === 'pool') {
    removeFieldEverywhere(field, null);
    if (state.chart.x === field) state.chart.x = null;
    if (state.chart.series === field) state.chart.series = null;
    if (state.chart.value === field) state.chart.value = null;
    refreshAll();
    return;
  }
  if (zoneKey === 'rows') {
    removeFieldEverywhere(field, 'rows');
    if (!state.pivot.rows.includes(field)) state.pivot.rows.push(field);
  } else if (zoneKey === 'cols') {
    removeFieldEverywhere(field, 'cols');
    if (!state.pivot.cols.includes(field)) state.pivot.cols.push(field);
  } else if (zoneKey === 'filters') {
    removeFieldEverywhere(field, 'filters');
    if (!state.pivot.filters.includes(field)) {
      state.pivot.filters.push(field);
      initFilterValues(field);
    }
  } else if (zoneKey === 'values') {
    if (!state.pivot.values.find(v => v.field === field)) {
      const defAgg = state.fieldTypes[field] === 'num' ? 'sum' : 'count';
      state.pivot.values.push({ field, agg: defAgg });
    }
  } else if (zoneKey === 'chartX') {
    state.chart.x = field;
  } else if (zoneKey === 'chartSeries') {
    state.chart.series = field;
  } else if (zoneKey === 'chartValue') {
    state.chart.value = field;
  }

  refreshAll();
}

function initFilterValues(field) {
  const distinct = getDistinctValues(field);
  if (!state._fieldDistinct) state._fieldDistinct = {};
  state._fieldDistinct[field] = new Set(distinct.map(String));
  if (!state.filterSelected[field]) {
    state.filterSelected[field] = new Set(distinct);
  }
}

let currentEditingFilterField = null;

function openFilterModal(field) {
  currentEditingFilterField = field;
  const modal = document.getElementById('filterModal');
  const title = document.getElementById('filterModalTitle');
  const sub = document.getElementById('filterModalSub');
  const listContainer = document.getElementById('filterValuesList');
  const searchInput = document.getElementById('filterValSearch');
  const counter = document.getElementById('filterSelectionCounter');

  if (!modal || !listContainer) return;

  title.textContent = `Filtrar: ${field}`;
  sub.textContent = `Selecione os valores permitidos para o campo "${field}".`;
  searchInput.value = '';

  const distinct = getDistinctValues(field);

  let currentSelected = new Set();
  const previous = state.filterSelected[field];
  if (previous) {
    distinct.forEach((v) => {
      if (selectionHas(previous, v)) currentSelected.add(v);
    });
  } else {
    distinct.forEach((v) => currentSelected.add(v));
  }

  function renderValueItems() {
    listContainer.innerHTML = '';
    const query = (searchInput.value || '').trim().toLowerCase();
    const visibleVals = distinct.filter(v => String(v).toLowerCase().includes(query));

    if (!visibleVals.length) {
      listContainer.innerHTML = '<div style="font-size:0.75rem; color:var(--color-text-muted); padding:0.5rem;">Nenhum valor encontrado.</div>';
      counter.textContent = `${currentSelected.size}/${distinct.length} Selecionados`;
      return;
    }

    visibleVals.forEach((val) => {
      const item = document.createElement('label');
      item.className = 'filter-value-item';

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = selectionHas(currentSelected, val);
      cb.addEventListener('change', () => {
        if (cb.checked) {
          currentSelected.add(val);
        } else {
          currentSelected.delete(val);
          currentSelected.delete(String(val));
        }
        counter.textContent = `${currentSelected.size}/${distinct.length} Selecionados`;
      });

      const txt = document.createElement('span');
      txt.textContent = String(val);

      item.appendChild(cb);
      item.appendChild(txt);
      listContainer.appendChild(item);
    });

    counter.textContent = `${currentSelected.size}/${distinct.length} Selecionados`;
  }

  searchInput.oninput = renderValueItems;

  const btnSelectAll = document.getElementById('btnSelectAllFilter');
  if (btnSelectAll) {
    btnSelectAll.onclick = () => {
      currentSelected = new Set(distinct);
      renderValueItems();
    };
  }

  const btnDeselectAll = document.getElementById('btnDeselectAllFilter');
  if (btnDeselectAll) {
    btnDeselectAll.onclick = () => {
      currentSelected.clear();
      renderValueItems();
    };
  }

  const btnApply = document.getElementById('btnApplyFilter');
  if (btnApply) {
    btnApply.onclick = () => {
      if (!state._fieldDistinct) state._fieldDistinct = {};
      state._fieldDistinct[field] = new Set(distinct.map(String));

      state.filterSelected[field] = new Set(currentSelected);
      const isInDimension = state.pivot.rows.includes(field) || state.pivot.cols.includes(field);
      if (!isInDimension && !state.pivot.filters.includes(field)) {
        state.pivot.filters.push(field);
      }
      modal.classList.add('hidden');
      refreshAll();
    };
  }

  renderValueItems();
  modal.classList.remove('hidden');
}

const closeFilterBtn = document.getElementById('closeFilterBtn');
if (closeFilterBtn) {
  closeFilterBtn.addEventListener('click', () => {
    const modal = document.getElementById('filterModal');
    if (modal) modal.classList.add('hidden');
  });
}

// RENDER DROP ZONE CHIPS
function renderAllZones() {
  renderSimpleZone('zoneRows', state.pivot.rows, 'rows');
  renderSimpleZone('zoneCols', state.pivot.cols, 'cols');
  renderFiltersZone();
  renderValuesZone();
  renderChartFieldZone('zoneChartX', state.chart.x, 'chartX');
  renderChartFieldZone('zoneChartSeries', state.chart.series, 'chartSeries');
  renderChartFieldZone('zoneChartValue', state.chart.value, 'chartValue');
}

function renderSimpleZone(elId, arr, kind) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = '';
  arr.forEach((field) => {
    const chip = document.createElement('div');
    chip.className = 'zone-chip';
    chip.draggable = true;
    
    const labelSpan = document.createElement('span');
    labelSpan.textContent = getFieldLabel(field);
    chip.appendChild(labelSpan);

    // Edit alias button
    const editBtn = document.createElement('span');
    editBtn.style.cursor = 'pointer';
    editBtn.style.fontSize = '0.7rem';
    editBtn.innerHTML = '✏️';
    editBtn.title = `Renomear rótulo de "${field}"`;
    editBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const current = getFieldLabel(field);
      const newAlias = prompt(`Renomear rótulo para o campo "${field}":`, current);
      if (newAlias && newAlias.trim()) {
        state.fieldAliases[field] = newAlias.trim();
        refreshAll();
      }
    });
    chip.appendChild(editBtn);

    // Filter button for row/col dimensions
    const filterBtn = document.createElement('span');
    filterBtn.className = 'filter-btn-pill';
    filterBtn.innerHTML = '▼';
    filterBtn.title = `Filtro e Ordenação Excel para "${getFieldLabel(field)}"`;
    filterBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      openExcelHeaderMenu(field);
    });
    chip.appendChild(filterBtn);

    chip.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', JSON.stringify({ field, from: kind }));
    });

    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = '✕';
    removeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      removeFrom(arr, field);
      refreshAll();
    });
    chip.appendChild(removeBtn);
    el.appendChild(chip);
  });
}

function renderFiltersZone() {
  const el = document.getElementById('zoneFilters');
  if (!el) return;
  el.innerHTML = '';
  state.pivot.filters.forEach((field) => {
    const chip = document.createElement('div');
    chip.className = 'zone-chip';
    chip.draggable = true;

    const distinct = getDistinctValues(field);
    const selectedSet = state.filterSelected[field] || new Set(distinct);
    const selCount = distinct.filter((v) => selectionHas(selectedSet, v)).length;
    const totalCount = distinct.length;
    const statusTxt = selCount === totalCount ? 'Todos' : `${selCount}/${totalCount}`;

    chip.innerHTML = `<span>${field}</span> <span class="filter-status-tag" title="Clique para editar valores permitidos">${statusTxt} 🔍</span>`;

    chip.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', JSON.stringify({ field, from: 'filters' }));
    });

    const tagEl = chip.querySelector('.filter-status-tag');
    if (tagEl) {
      tagEl.addEventListener('click', (e) => {
        e.stopPropagation();
        openFilterModal(field);
      });
    }

    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = '✕';
    removeBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      removeFrom(state.pivot.filters, field);
      delete state.filterSelected[field];
      refreshAll();
    });

    chip.appendChild(removeBtn);
    el.appendChild(chip);
  });
}

function renderValuesZone() {
  const el = document.getElementById('zoneValues');
  if (!el) return;
  el.innerHTML = '';
  state.pivot.values.forEach((v, idx) => {
    const chip = document.createElement('div');
    chip.className = 'zone-chip';
    chip.draggable = true;
    chip.innerHTML = `<span>${v.field}</span>`;
    chip.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', JSON.stringify({ field: v.field, from: 'values' }));
    });

    const aggSel = document.createElement('select');
    Object.keys(AGGS).forEach((k) => {
      const o = document.createElement('option');
      o.value = k;
      o.textContent = AGGS[k].label;
      if (k === v.agg) o.selected = true;
      aggSel.appendChild(o);
    });
    aggSel.addEventListener('change', () => {
      v.agg = aggSel.value;
      renderMatrix();
    });

    chip.appendChild(aggSel);
    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove';
    removeBtn.textContent = '✕';
    removeBtn.addEventListener('click', () => {
      state.pivot.values.splice(idx, 1);
      refreshAll();
    });
    chip.appendChild(removeBtn);
    el.appendChild(chip);
  });
}

function renderChartFieldZone(elId, field, kind) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = '';
  if (!field) return;
  const chip = document.createElement('div');
  chip.className = 'zone-chip';
  chip.draggable = true;
  chip.innerHTML = `<span>${field}</span>`;
  chip.addEventListener('dragstart', (e) => {
    e.dataTransfer.setData('text/plain', JSON.stringify({ field, from: kind }));
  });
  const removeBtn = document.createElement('button');
  removeBtn.className = 'btn-remove';
  removeBtn.textContent = '✕';
  removeBtn.addEventListener('click', () => {
    if (kind === 'chartX') state.chart.x = null;
    if (kind === 'chartSeries') state.chart.series = null;
    if (kind === 'chartValue') state.chart.value = null;
    refreshAll();
  });
  chip.appendChild(removeBtn);
  el.appendChild(chip);
}

// FILTER DATA RESTRAINER
// Applies ALL active filterSelected restrictions — for global filters (pivot.filters)
// AND for dimension fields (rows/cols) filtered via the Excel Header menu.
function getFilteredData() {
  const activeFilters = Object.keys(state.filterSelected).filter((f) => {
    const sel = state.filterSelected[f];
    if (!sel || sel.size === 0) return true;
    const distinct = getDistinctValues(f);
    // Sem restrição efetiva se todos os valores distintos estão selecionados
    return !distinct.every((v) => selectionHas(sel, v));
  });

  if (!activeFilters.length) return state.data;

  return state.data.filter((row) => {
    for (const f of activeFilters) {
      const sel = state.filterSelected[f];
      if (!sel) continue;
      if (sel.size === 0) return false;

      const normVal = normalizeDimValue(row[f]);
      if (!selectionHas(sel, normVal)) return false;
    }
    return true;
  });
}


// CALENDAR & MONTH SORTING LOGIC FOR DIMENSION COMBOS
function sortCombos(combos, fields) {
  combos.sort((a, b) => {
    for (let i = 0; i < a.length; i++) {
      const fieldName = fields[i];
      const sortDir = state.fieldSortDirs[fieldName] || 'asc';
      const factor = sortDir === 'desc' ? -1 : 1;
      const av = a[i], bv = b[i];
      if (av === bv) continue;

      // Check if field is month-based
      if (isMonthField(fieldName)) {
        const mIdxA = getMonthIndex(av);
        const mIdxB = getMonthIndex(bv);
        if (mIdxA !== 99 || mIdxB !== 99) {
          if (mIdxA !== mIdxB) return (mIdxA - mIdxB) * factor;
        }
      }

      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * factor;
      return ('' + av).localeCompare(('' + bv), 'pt-BR', { numeric: true }) * factor;
    }
    return 0;
  });
}

function comboKey(combo) {
  return combo.map(v => String(v)).join('\u0001');
}

function uniqueCombos(data, fields) {
  if (fields.length === 0) return [[]];
  const map = new Map();
  data.forEach((row) => {
    const combo = fields.map(f => normalizeDimValue(row[f]));
    map.set(comboKey(combo), combo);
  });
  const combos = Array.from(map.values());
  sortCombos(combos, fields);
  return combos;
}

function buildPivotModel() {
  const data = getFilteredData();
  const rowFields = state.pivot.rows;
  const colFields = state.pivot.cols;
  const values = state.pivot.values.length ? state.pivot.values : [{ field: null, agg: 'count' }];

  const rowCombos = uniqueCombos(data, rowFields);
  const colCombos = uniqueCombos(data, colFields);

  const buckets = new Map();
  data.forEach((row) => {
    const rk = rowFields.map(f => normalizeDimValue(row[f])).join('\u0001');
    const ck = colFields.map(f => normalizeDimValue(row[f])).join('\u0001');
    const key = rk + '\u0002' + ck;
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(row);
  });

  function cellAgg(rowCombo, colCombo, valDef) {
    const rk = rowCombo.join('\u0001');
    const ck = colCombo.join('\u0001');
    const bucket = buckets.get(rk + '\u0002' + ck) || [];
    if (!valDef.field) return AGGS.count.fn(bucket.map(() => 1));
    const nums = bucket.map((r) => {
      const v = r[valDef.field];
      return typeof v === 'number' ? v : (parseFloat(v) || 0);
    });
    return AGGS[valDef.agg] ? AGGS[valDef.agg].fn(nums) : AGGS.sum.fn(nums);
  }

  return { data, rowFields, colFields, values, rowCombos, colCombos, cellAgg };
}

// EXCEL HEADER MENU POPOVER (SORTING, FILTERING & ALIAS RENAMING)
function openExcelHeaderMenu(field) {
  const modal = document.getElementById('excelHeaderMenuModal');
  const title = document.getElementById('excelMenuFieldTitle');
  const sub = document.getElementById('excelMenuFieldSub');
  const renameInput = document.getElementById('excelRenameInput');
  const btnSaveAlias = document.getElementById('btnSaveAlias');
  const btnSortAsc = document.getElementById('btnSortAsc');
  const btnSortDesc = document.getElementById('btnSortDesc');
  const searchInput = document.getElementById('excelFilterSearch');
  const listContainer = document.getElementById('excelFilterValuesList');
  const counter = document.getElementById('excelFilterCounter');
  const btnApply = document.getElementById('btnApplyExcelMenu');

  if (!modal || !listContainer) return;

  const currentLabel = getFieldLabel(field);
  title.textContent = `Filtro de Campo: ${currentLabel}`;
  sub.textContent = field !== currentLabel ? `Campo original: "${field}"` : 'Opções de ordenação, renomeação e filtragem estilo Excel.';
  renameInput.value = currentLabel;

  let currentSort = state.fieldSortDirs[field] || 'asc';

  function updateSortBtnStyles() {
    if (currentSort === 'asc') {
      btnSortAsc.style.background = 'var(--color-primary)';
      btnSortAsc.style.color = '#fff';
      btnSortDesc.style.background = 'transparent';
      btnSortDesc.style.color = 'var(--color-text-main)';
    } else {
      btnSortDesc.style.background = 'var(--color-primary)';
      btnSortDesc.style.color = '#fff';
      btnSortAsc.style.background = 'transparent';
      btnSortAsc.style.color = 'var(--color-text-main)';
    }
  }
  updateSortBtnStyles();

  btnSortAsc.onclick = () => {
    currentSort = 'asc';
    updateSortBtnStyles();
  };

  btnSortDesc.onclick = () => {
    currentSort = 'desc';
    updateSortBtnStyles();
  };

  const distinct = getDistinctValues(field);

  let currentSelected = new Set();
  const previous = state.filterSelected[field];
  if (previous) {
    distinct.forEach((v) => {
      if (selectionHas(previous, v)) currentSelected.add(v);
    });
  } else {
    distinct.forEach((v) => currentSelected.add(v));
  }

  function renderValueItems() {
    listContainer.innerHTML = '';
    const query = (searchInput.value || '').trim().toLowerCase();
    const visibleVals = distinct.filter(v => String(v).toLowerCase().includes(query));

    if (!visibleVals.length) {
      listContainer.innerHTML = '<div style="font-size:0.75rem; color:var(--color-text-muted); padding:0.5rem;">Nenhum item encontrado.</div>';
      counter.textContent = `${currentSelected.size}/${distinct.length}`;
      return;
    }

    visibleVals.forEach((val) => {
      const item = document.createElement('label');
      item.className = 'filter-value-item';

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = selectionHas(currentSelected, val);
      cb.addEventListener('change', () => {
        if (cb.checked) currentSelected.add(val);
        else {
          currentSelected.delete(val);
          currentSelected.delete(String(val));
        }
        counter.textContent = `${currentSelected.size}/${distinct.length}`;
      });

      const txt = document.createElement('span');
      txt.textContent = String(val);

      item.appendChild(cb);
      item.appendChild(txt);
      listContainer.appendChild(item);
    });

    counter.textContent = `${currentSelected.size}/${distinct.length}`;
  }

  searchInput.value = '';
  searchInput.oninput = renderValueItems;

  document.getElementById('btnExcelSelectAll').onclick = () => {
    currentSelected = new Set(distinct);
    renderValueItems();
  };

  document.getElementById('btnExcelDeselectAll').onclick = () => {
    currentSelected.clear();
    renderValueItems();
  };

  btnSaveAlias.onclick = () => {
    const newAlias = renameInput.value.trim();
    if (newAlias) {
      state.fieldAliases[field] = newAlias;
      title.textContent = `Filtro de Campo: ${newAlias}`;
      refreshAll();
    }
  };

  btnApply.onclick = () => {
    const newAlias = renameInput.value.trim();
    if (newAlias) {
      state.fieldAliases[field] = newAlias;
    }

    state.fieldSortDirs[field] = currentSort;

    if (!state._fieldDistinct) state._fieldDistinct = {};
    state._fieldDistinct[field] = new Set(distinct.map(String));

    state.filterSelected[field] = new Set(currentSelected);

    const isInDimension = state.pivot.rows.includes(field) || state.pivot.cols.includes(field);
    if (!isInDimension && !state.pivot.filters.includes(field) && currentSelected.size < distinct.length) {
      state.pivot.filters.push(field);
    }

    modal.classList.add('hidden');
    refreshAll();
  };

  renderValueItems();
  modal.classList.remove('hidden');
}

const closeExcelMenuBtn = document.getElementById('closeExcelMenuBtn');
if (closeExcelMenuBtn) {
  closeExcelMenuBtn.addEventListener('click', () => {
    const modal = document.getElementById('excelHeaderMenuModal');
    if (modal) modal.classList.add('hidden');
  });
}

// MATRIX RENDERING ENGINE WITH CASCADE TREE COLLAPSE & EXCEL HEADERS
function renderMatrix() {
  const host = document.getElementById('matrixHost');
  if (!host) return;

  if (!state.data.length) {
    host.innerHTML = `
      <div class="empty-state-hero">
        <div class="hero-icon">📊</div>
        <h3>Matriz Bi(Cubo) Pronta</h3>
        <p>Carregue os dados de demonstração ou um arquivo Excel para explorar a matriz.</p>
        <button id="btnQuickDemoHero2" class="btn btn-primary">Carregar Dados de Exemplo</button>
      </div>`;
    const btn2 = document.getElementById('btnQuickDemoHero2');
    if (btn2) btn2.addEventListener('click', () => loadDataset(generateDemoDataset(), 'Vendas Empresariais (Demo)'));
    state.lastMatrixAOA = null;
    updateExportButtons();
    return;
  }

  const m = buildPivotModel();
  const { rowFields, colFields, values, rowCombos, colCombos, cellAgg } = m;

  const nRowLevels = Math.max(rowFields.length, 1);
  const nColLevels = colFields.length;
  const nVals = values.length;

  // TABLE HEADERS WITH EXCEL DROPDOWNS
  const theadRows = [];
  for (let lvl = 0; lvl < nColLevels; lvl++) {
    const tr = document.createElement('tr');
    if (lvl === 0) {
      const corner = document.createElement('th');
      corner.colSpan = nRowLevels;
      corner.rowSpan = nColLevels + 1;

      if (rowFields.length) {
        const cornerWrap = document.createElement('div');
        cornerWrap.style.display = 'flex';
        cornerWrap.style.alignItems = 'center';
        cornerWrap.style.justifyContent = 'center';
        cornerWrap.style.gap = '0.4rem';
        cornerWrap.style.flexWrap = 'wrap';

        rowFields.forEach((f, idx) => {
          if (idx > 0) {
            const sep = document.createElement('span');
            sep.textContent = '➔';
            sep.style.color = '#94A3B8';
            cornerWrap.appendChild(sep);
          }

          const btn = document.createElement('button');
          btn.className = 'excel-header-btn';
          btn.title = `Filtro e Ordenação Excel para "${getFieldLabel(f)}"`;
          btn.innerHTML = `<span>${getFieldLabel(f)}</span><span class="excel-arrow">▼</span>`;
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            openExcelHeaderMenu(f);
          });
          cornerWrap.appendChild(btn);
        });

        corner.appendChild(cornerWrap);
      } else {
        corner.textContent = '';
      }

      tr.appendChild(corner);
    }

    const colField = colFields[lvl];
    colCombos.forEach((cc) => {
      const th = document.createElement('th');
      th.colSpan = nVals;
      const colVal = String(cc[lvl] || '(vazio)');

      const btn = document.createElement('button');
      btn.className = 'excel-header-btn';
      btn.title = `Filtro Excel para ${colField ? getFieldLabel(colField) : ''}`;
      btn.innerHTML = `<span>${colVal}</span><span class="excel-arrow">▼</span>`;
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (colField) {
          openExcelHeaderMenu(colField);
        }
      });

      th.appendChild(btn);
      tr.appendChild(th);
    });

    if (lvl === 0) {
      const th = document.createElement('th');
      th.rowSpan = nColLevels + 1;
      th.className = 'total-col';
      th.textContent = 'Total Geral';
      tr.appendChild(th);
    }
    theadRows.push(tr);
  }

  // Value Sub-headers Row
  const valSubTr = document.createElement('tr');
  if (nColLevels === 0) {
    const corner = document.createElement('th');
    corner.colSpan = nRowLevels;
    if (rowFields.length) {
      const cornerWrap = document.createElement('div');
      cornerWrap.style.display = 'flex';
      cornerWrap.style.alignItems = 'center';
      cornerWrap.style.justifyContent = 'center';
      cornerWrap.style.gap = '0.4rem';
      cornerWrap.style.flexWrap = 'wrap';

      rowFields.forEach((f, idx) => {
        if (idx > 0) {
          const sep = document.createElement('span');
          sep.textContent = '➔';
          sep.style.color = '#94A3B8';
          cornerWrap.appendChild(sep);
        }

        const btn = document.createElement('button');
        btn.className = 'excel-header-btn';
        btn.title = `Filtro e Ordenação Excel para "${getFieldLabel(f)}"`;
        btn.innerHTML = `<span>${getFieldLabel(f)}</span><span class="excel-arrow">▼</span>`;
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          openExcelHeaderMenu(f);
        });
        cornerWrap.appendChild(btn);
      });

      corner.appendChild(cornerWrap);
    } else {
      corner.textContent = '';
    }
    valSubTr.appendChild(corner);
  }

  colCombos.forEach(() => {
    values.forEach((v) => {
      const th = document.createElement('th');
      const valLabel = v.field ? `${AGGS[v.agg].label} ${getFieldLabel(v.field)}` : 'Contagem';

      const span = document.createElement('span');
      span.textContent = valLabel;
      th.appendChild(span);

      if (v.field) {
        const editBtn = document.createElement('span');
        editBtn.style.cursor = 'pointer';
        editBtn.style.marginLeft = '0.35rem';
        editBtn.style.fontSize = '0.7rem';
        editBtn.innerHTML = '✏️';
        editBtn.title = `Renomear rótulo de "${v.field}"`;
        editBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          const current = getFieldLabel(v.field);
          const newAlias = prompt(`Renomear rótulo para "${v.field}":`, current);
          if (newAlias && newAlias.trim()) {
            state.fieldAliases[v.field] = newAlias.trim();
            refreshAll();
          }
        });
        th.appendChild(editBtn);
      }

      valSubTr.appendChild(th);
    });
  });
  if (nColLevels === 0) {
    const th = document.createElement('th');
    th.className = 'total-col';
    th.textContent = 'Total Geral';
    valSubTr.appendChild(th);
  }
  theadRows.push(valSubTr);

  // TABLE BODY WITH EXCEL / GOOGLE SHEETS GROUPING & SUBTOTALS
  const tbodyRows = [];
  const treeNodes = buildRowHierarchyTree(m.data, rowFields);

  if (rowFields.length > 0) {
    const generatedRows = generateTableRowsFromTree(treeNodes, rowFields, colFields, colCombos, values, cellAgg, m);
    generatedRows.forEach(r => tbodyRows.push(r));
  } else {
    // Single Total Row if no row fields specified
    const tr = document.createElement('tr');
    const th = document.createElement('th');
    th.textContent = 'Total';
    tr.appendChild(th);

    colCombos.forEach((cc) => {
      values.forEach((v) => {
        const val = cellAgg([], cc, v);
        const td = document.createElement('td');
        td.className = 'val-cell';
        td.textContent = fmtNum(val);
        tr.appendChild(td);
      });
    });

    values.forEach((v) => {
      const nums = m.data.map(r => typeof r[v.field] === 'number' ? r[v.field] : (parseFloat(r[v.field]) || 0));
      const val = !v.field ? m.data.length : AGGS[v.agg].fn(nums);
      const td = document.createElement('td');
      td.className = 'val-cell total-col';
      td.textContent = fmtNum(val);
      tr.appendChild(td);
    });

    tbodyRows.push(tr);
  }

  // GRAND TOTAL ROW
  const totalTr = document.createElement('tr');
  totalTr.className = 'total-row';
  const th0 = document.createElement('th');
  th0.colSpan = nRowLevels;
  th0.textContent = 'Total Geral';
  totalTr.appendChild(th0);

  colCombos.forEach((cc) => {
    values.forEach((v) => {
      const allInCol = m.data.filter((r) => colFields.length ? colFields.every((f, i) => dimValuesMatch(r[f], cc[i])) : true);
      const nums = allInCol.map(r => typeof r[v.field] === 'number' ? r[v.field] : (parseFloat(r[v.field]) || 0));
      const val = !v.field ? allInCol.length : AGGS[v.agg].fn(nums);
      const td = document.createElement('td');
      td.className = 'val-cell';
      td.textContent = fmtNum(val);
      totalTr.appendChild(td);
    });
  });

  values.forEach((v) => {
    const nums = m.data.map(r => typeof r[v.field] === 'number' ? r[v.field] : (parseFloat(r[v.field]) || 0));
    const val = !v.field ? m.data.length : AGGS[v.agg].fn(nums);
    const td = document.createElement('td');
    td.className = 'val-cell total-col';
    td.textContent = fmtNum(val);
    totalTr.appendChild(td);
  });

  // BUILD FINAL HTML TABLE
  const table = document.createElement('table');
  table.className = 'matrix-table';
  const thead = document.createElement('thead');
  theadRows.forEach(r => thead.appendChild(r));
  table.appendChild(thead);
  const tbody = document.createElement('tbody');
  tbodyRows.forEach(r => tbody.appendChild(r));
  tbody.appendChild(totalTr);
  table.appendChild(tbody);

  host.innerHTML = '';
  host.appendChild(table);

  const totalRowCount = rowCombos.length;
  document.getElementById('matrixStatsLabel').textContent = `${totalRowCount} grupos × ${colCombos.length} colunas`;

  buildMatrixAOA(m, rowCombos.length ? rowCombos : [['Total']]);
  updateExportButtons();
}

// TREE HIERARCHY BUILDER FOR ROW GROUPING
function buildRowHierarchyTree(data, rowFields) {
  if (!rowFields.length) return [];

  function buildLevel(rows, levelIndex, parentPathKey) {
    if (levelIndex >= rowFields.length) return [];

    const field = rowFields[levelIndex];
    const map = new Map();

    rows.forEach((r) => {
      const v = normalizeDimValue(r[field]);
      if (!map.has(v)) map.set(v, []);
      map.get(v).push(r);
    });

    const sortedCombos = Array.from(map.keys()).map((v) => [v]);
    sortCombos(sortedCombos, [field]);
    const values = sortedCombos.map((c) => c[0]);

    return values.map((val) => {
      const groupRows = map.get(val);
      const currentKey = parentPathKey ? `${parentPathKey}\u0001${val}` : String(val);
      const children = levelIndex + 1 < rowFields.length 
        ? buildLevel(groupRows, levelIndex + 1, currentKey) 
        : [];

      return {
        field,
        value: val,
        level: levelIndex,
        key: currentKey,
        rows: groupRows,
        children
      };
    });
  }

  return buildLevel(data, 0, '');
}

// RECURSIVE TABLE ROWS GENERATOR FOR GROUPING & SUBTOTALS
function generateTableRowsFromTree(nodes, rowFields, colFields, colCombos, values, cellAgg, m) {
  const trs = [];

  function traverseNode(node) {
    const isCollapsed = state.collapsedNodes.has(node.key);
    const hasChildren = node.children && node.children.length > 0;

    if (hasChildren && isCollapsed) {
      // COLLAPSED GROUP SUMMARY ROW (Excel / Google Sheets style)
      const tr = document.createElement('tr');
      tr.className = 'subtotal-row collapsed-group';

      const th = document.createElement('th');
      th.colSpan = rowFields.length - node.level;
      th.style.paddingLeft = `${node.level * 16 + 8}px`;

      const toggleBtn = document.createElement('span');
      toggleBtn.className = 'tree-toggle-btn';
      toggleBtn.textContent = '+';
      toggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        state.collapsedNodes.delete(node.key);
        renderMatrix();
      });
      th.appendChild(toggleBtn);

      const labelTxt = document.createTextNode(` [+] Total ${node.value}`);
      th.appendChild(labelTxt);
      tr.appendChild(th);

      // Compute subtotal metrics for collapsed group
      colCombos.forEach((cc) => {
        values.forEach((v) => {
          const val = aggregateBucket(node.rows, colFields, cc, v);
          const td = document.createElement('td');
          td.className = 'val-cell';
          td.textContent = fmtNum(val);
          tr.appendChild(td);
        });
      });

      values.forEach((v) => {
        const nums = node.rows.map(r => typeof r[v.field] === 'number' ? r[v.field] : (parseFloat(r[v.field]) || 0));
        const val = !v.field ? node.rows.length : AGGS[v.agg].fn(nums);
        const td = document.createElement('td');
        td.className = 'val-cell total-col';
        td.textContent = fmtNum(val);
        tr.appendChild(td);
      });

      trs.push(tr);
      return;
    }

    if (hasChildren) {
      // EXPANDED GROUP PARENT HEADER (In compact mode)
      if (state.layoutMode === 'compact') {
        const tr = document.createElement('tr');
        tr.className = 'group-parent-row';

        const th = document.createElement('th');
        th.colSpan = rowFields.length;
        th.style.paddingLeft = `${node.level * 16 + 8}px`;

        const toggleBtn = document.createElement('span');
        toggleBtn.className = 'tree-toggle-btn';
        toggleBtn.textContent = '-';
        toggleBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          state.collapsedNodes.add(node.key);
          renderMatrix();
        });
        th.appendChild(toggleBtn);

        const labelTxt = document.createTextNode(` [-] ${node.value}`);
        th.appendChild(labelTxt);
        tr.appendChild(th);

        colCombos.forEach(() => {
          values.forEach(() => {
            const td = document.createElement('td');
            td.className = 'val-cell';
            td.textContent = '';
            tr.appendChild(td);
          });
        });
        values.forEach(() => {
          const td = document.createElement('td');
          td.className = 'val-cell total-col';
          td.textContent = '';
          tr.appendChild(td);
        });

        trs.push(tr);
      }

      // Process children
      node.children.forEach((child) => {
        traverseNode(child);
      });

      // SUBTOTAL ROW FOR EXPANDED GROUP (Excel / Google Sheets style)
      if (state.showSubtotals) {
        const subTr = document.createElement('tr');
        subTr.className = 'subtotal-row';

        const th = document.createElement('th');
        th.colSpan = rowFields.length - node.level;
        th.style.paddingLeft = `${node.level * 16 + 8}px`;
        th.textContent = `∑ Total ${node.value}`;
        subTr.appendChild(th);

        colCombos.forEach((cc) => {
          values.forEach((v) => {
            const val = aggregateBucket(node.rows, colFields, cc, v);
            const td = document.createElement('td');
            td.className = 'val-cell';
            td.textContent = fmtNum(val);
            subTr.appendChild(td);
          });
        });

        values.forEach((v) => {
          const nums = node.rows.map(r => typeof r[v.field] === 'number' ? r[v.field] : (parseFloat(r[v.field]) || 0));
          const val = !v.field ? node.rows.length : AGGS[v.agg].fn(nums);
          const td = document.createElement('td');
          td.className = 'val-cell total-col';
          td.textContent = fmtNum(val);
          subTr.appendChild(td);
        });

        trs.push(subTr);
      }
    } else {
      // LEAF ROW
      const tr = document.createElement('tr');
      const pathParts = node.key.split('\u0001');

      if (state.layoutMode === 'tabular') {
        pathParts.forEach((val, lvlIdx) => {
          const td = document.createElement('td');
          td.style.fontWeight = lvlIdx === 0 ? '600' : '400';
          td.style.paddingLeft = `${lvlIdx * 16 + 8}px`;

          const currentKey = pathParts.slice(0, lvlIdx + 1).join('\u0001');
          const isNodeParent = rowFields.length > lvlIdx + 1;

          if (isNodeParent) {
            const toggleBtn = document.createElement('span');
            toggleBtn.className = 'tree-toggle-btn';
            toggleBtn.textContent = state.collapsedNodes.has(currentKey) ? '+' : '-';
            toggleBtn.addEventListener('click', (e) => {
              e.stopPropagation();
              if (state.collapsedNodes.has(currentKey)) state.collapsedNodes.delete(currentKey);
              else state.collapsedNodes.add(currentKey);
              renderMatrix();
            });
            td.appendChild(toggleBtn);
          }

          td.appendChild(document.createTextNode(String(val)));
          tr.appendChild(td);
        });
      } else {
        // Compact mode leaf row
        const th = document.createElement('th');
        th.colSpan = rowFields.length;
        th.style.paddingLeft = `${(node.level + 1) * 16 + 8}px`;
        th.textContent = String(node.value);
        tr.appendChild(th);
      }

      // Fill leaf metric values
      colCombos.forEach((cc) => {
        values.forEach((v) => {
          const val = aggregateBucket(node.rows, colFields, cc, v);
          const td = document.createElement('td');
          td.className = 'val-cell';
          td.textContent = fmtNum(val);
          tr.appendChild(td);
        });
      });

      // Leaf row grand total
      values.forEach((v) => {
        const nums = node.rows.map(r => typeof r[v.field] === 'number' ? r[v.field] : (parseFloat(r[v.field]) || 0));
        const val = !v.field ? node.rows.length : AGGS[v.agg].fn(nums);
        const td = document.createElement('td');
        td.className = 'val-cell total-col';
        td.textContent = fmtNum(val);
        tr.appendChild(td);
      });

      trs.push(tr);
    }
  }

  nodes.forEach((n) => traverseNode(n));
  return trs;
}

function aggregateBucket(rows, colFields, colCombo, valDef) {
  const matched = rows.filter((r) => {
    if (!colFields.length) return true;
    return colFields.every((f, i) => dimValuesMatch(r[f], colCombo[i]));
  });

  if (!valDef.field) return AGGS.count.fn(matched.map(() => 1));
  const nums = matched.map(r => typeof r[valDef.field] === 'number' ? r[valDef.field] : (parseFloat(r[valDef.field]) || 0));
  return AGGS[valDef.agg] ? AGGS[valDef.agg].fn(nums) : AGGS.sum.fn(nums);
}

// TREE EXPAND/COLLAPSE TOOLBAR CONTROLS
const btnExpandAll = document.getElementById('btnExpandAll');
if (btnExpandAll) {
  btnExpandAll.addEventListener('click', () => {
    state.collapsedNodes.clear();
    renderMatrix();
  });
}

const btnCollapseAll = document.getElementById('btnCollapseAll');
if (btnCollapseAll) {
  btnCollapseAll.addEventListener('click', () => {
    state.pivot.rows.forEach((_, idx) => {
      state.data.forEach((r) => {
        const key = state.pivot.rows.slice(0, idx + 1).map(f => normalizeDimValue(r[f])).join('\u0001');
        state.collapsedNodes.add(key);
      });
    });
    renderMatrix();
  });
}

const btnToggleSubtotals = document.getElementById('btnToggleSubtotals');
if (btnToggleSubtotals) {
  btnToggleSubtotals.addEventListener('click', () => {
    state.showSubtotals = !state.showSubtotals;
    btnToggleSubtotals.innerHTML = `<span>∑ Subtotais: ${state.showSubtotals ? 'Sim' : 'Não'}</span>`;
    renderMatrix();
  });
}

const btnToggleLayoutMode = document.getElementById('btnToggleLayoutMode');
if (btnToggleLayoutMode) {
  btnToggleLayoutMode.addEventListener('click', () => {
    state.layoutMode = state.layoutMode === 'tabular' ? 'compact' : 'tabular';
    btnToggleLayoutMode.innerHTML = `<span>📐 Layout: ${state.layoutMode === 'tabular' ? 'Tabular' : 'Compacto'}</span>`;
    renderMatrix();
  });
}

// ARRAY-OF-ARRAYS GENERATOR FOR XLSX & PDF EXPORTS
function buildMatrixAOA(m, effectiveRowCombos) {
  const { rowFields, colFields, values, colCombos, cellAgg } = m;
  const header = [...(rowFields.length ? rowFields : ['(sem linha)'])];
  
  colCombos.forEach((cc) => {
    const colLbl = colFields.length ? cc.join(' / ') : 'Total';
    values.forEach(v => header.push(`${colLbl} — ${v.field ? AGGS[v.agg].label + ' ' + v.field : 'Contagem'}`));
  });
  values.forEach(v => header.push(`Total Geral — ${v.field ? AGGS[v.agg].label + ' ' + v.field : 'Contagem'}`));

  const aoa = [header];

  effectiveRowCombos.forEach((rc) => {
    const line = [...rc];
    colCombos.forEach((cc) => {
      values.forEach((v) => {
        const val = rowFields.length ? cellAgg(rc, cc, v) : cellAgg([], cc, v);
        line.push(Math.round(val * 100) / 100);
      });
    });
    values.forEach((v) => {
      const allInRow = m.data.filter(r => rowFields.length ? rowFields.every((f, i) => dimValuesMatch(r[f], rc[i])) : true);
      const nums = allInRow.map(r => typeof r[v.field] === 'number' ? r[v.field] : (parseFloat(r[v.field]) || 0));
      const val = !v.field ? allInRow.length : AGGS[v.agg].fn(nums);
      line.push(Math.round(val * 100) / 100);
    });
    aoa.push(line);
  });

  state.lastMatrixAOA = aoa;
}

// CHART STUDIO ENGINE
function renderChart() {
  const area = document.getElementById('chartArea');
  if (!area) return;

  const { x, series, value, agg } = state.chart;
  if (!state.data.length || !x || !value) {
    area.innerHTML = `
      <div class="empty-state-hero">
        <div class="hero-icon">📈</div>
        <h3>Nenhum Gráfico Configurado</h3>
        <p>Arraste um campo para o Eixo X e um campo numérico para Valor no painel lateral.</p>
      </div>`;
    if (state.chartInstance) {
      state.chartInstance.destroy();
      state.chartInstance = null;
    }
    document.getElementById('btnExportChartPdf').disabled = true;
    document.getElementById('btnExportChartPng').disabled = true;
    return;
  }

  const data = getFilteredData();
  const aggFn = AGGS[agg] ? AGGS[agg].fn : AGGS.sum.fn;

  let xVals = Array.from(new Set(data.map(r => normalizeDimValue(r[x]))));

  if (isMonthField(x)) {
    xVals.sort((a, b) => getMonthIndex(a) - getMonthIndex(b));
  } else {
    xVals.sort((a, b) => ('' + a).localeCompare(('' + b), 'pt-BR', { numeric: true }));
  }

  const palette = ['#1E3A8A', '#10B981', '#D97706', '#EF4444', '#6366F1', '#8B5CF6', '#EC4899', '#06B6D4'];

  let datasets = [];
  if (series) {
    const sVals = Array.from(new Set(data.map(r => normalizeDimValue(r[series]))));
    sVals.forEach((sv, i) => {
      const arr = xVals.map((xv) => {
        const rows = data.filter(r => normalizeDimValue(r[x]) === xv && normalizeDimValue(r[series]) === sv);
        const nums = rows.map(r => typeof r[value] === 'number' ? r[value] : (parseFloat(r[value]) || 0));
        return nums.length ? aggFn(nums) : 0;
      });
      datasets.push({
        label: String(sv),
        data: arr,
        backgroundColor: palette[i % palette.length],
        borderColor: palette[i % palette.length]
      });
    });
  } else {
    const arr = xVals.map((xv) => {
      const rows = data.filter(r => normalizeDimValue(r[x]) === xv);
      const nums = rows.map(r => typeof r[value] === 'number' ? r[value] : (parseFloat(r[value]) || 0));
      return nums.length ? aggFn(nums) : 0;
    });
    datasets.push({
      label: `${AGGS[agg] ? AGGS[agg].label : 'Soma'} de ${value}`,
      data: arr,
      backgroundColor: xVals.map((_, i) => palette[i % palette.length]),
      borderColor: palette[0]
    });
  }

  area.innerHTML = '<canvas id="chartCanvas"></canvas>';
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (state.chartInstance) state.chartInstance.destroy();

  const type = document.getElementById('chartType').value;
  state.chartInstance = new Chart(ctx, {
    type: type === 'area' ? 'line' : (type === 'scatter' ? 'scatter' : type),
    data: {
      labels: xVals.map(String),
      datasets: datasets.map(d => type === 'area' ? { ...d, fill: true } : d)
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: !!series || type === 'pie' }
      },
      scales: type === 'pie' ? {} : { y: { beginAtZero: true } }
    }
  });

  document.getElementById('btnExportChartPdf').disabled = false;
  document.getElementById('btnExportChartPng').disabled = false;
}

document.getElementById('chartType').addEventListener('change', renderChart);

// TAB SWITCHER
document.querySelectorAll('.tab-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    const targetId = btn.dataset.tab;
    const targetContent = document.getElementById(targetId);
    if (targetContent) targetContent.classList.add('active');
  });
});

// =====================================================================
// FULLSCREEN MATRIX VIEW
// =====================================================================
const btnToggleFullscreen = document.getElementById('btnToggleFullscreen');

function enterFullscreen() {
  document.body.classList.add('fullscreen-matrix-active');
  if (btnToggleFullscreen) {
    btnToggleFullscreen.innerHTML = '<span>✕ Sair da Tela Cheia</span>';
    btnToggleFullscreen.title = 'Sair da visualização em tela cheia (Esc)';
  }
}

function exitFullscreen() {
  document.body.classList.remove('fullscreen-matrix-active');
  if (btnToggleFullscreen) {
    btnToggleFullscreen.innerHTML = '<span>⛶ Tela Cheia</span>';
    btnToggleFullscreen.title = 'Expandir a matriz para visualização em tela cheia (100% da tela)';
  }
}

if (btnToggleFullscreen) {
  btnToggleFullscreen.addEventListener('click', () => {
    if (document.body.classList.contains('fullscreen-matrix-active')) {
      exitFullscreen();
    } else {
      enterFullscreen();
    }
  });
}

// Exit fullscreen with Escape key
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && document.body.classList.contains('fullscreen-matrix-active')) {
    exitFullscreen();
  }
});

// =====================================================================
// MOBILE SIDEBAR DRAWER TOGGLE
// =====================================================================
const btnMobileToggleSidebar = document.getElementById('btnMobileToggleSidebar');
const sidebarEl = document.querySelector('.sidebar');

function openMobileSidebar() {
  if (sidebarEl) sidebarEl.classList.add('sidebar-drawer-open');
  // Create overlay if not exists
  if (!document.getElementById('sidebarOverlay')) {
    const overlay = document.createElement('div');
    overlay.id = 'sidebarOverlay';
    overlay.className = 'sidebar-overlay';
    overlay.addEventListener('click', closeMobileSidebar);
    document.body.appendChild(overlay);
  }
}

function closeMobileSidebar() {
  if (sidebarEl) sidebarEl.classList.remove('sidebar-drawer-open');
  const overlay = document.getElementById('sidebarOverlay');
  if (overlay) overlay.remove();
}

if (btnMobileToggleSidebar) {
  btnMobileToggleSidebar.addEventListener('click', () => {
    if (sidebarEl && sidebarEl.classList.contains('sidebar-drawer-open')) {
      closeMobileSidebar();
    } else {
      openMobileSidebar();
    }
  });
}

// EXPORT FUNCTIONALITIES
function updateExportButtons() {
  const hasData = !!state.lastMatrixAOA;
  document.getElementById('btnExportXlsx').disabled = !hasData;
  document.getElementById('btnExportXls').disabled = !hasData;
  document.getElementById('btnExportPdf').disabled = !hasData;
  document.getElementById('btnSaveCube').disabled = !state.data.length;
}

function exportMatrix(bookType) {
  if (!state.lastMatrixAOA) return;
  const ws = XLSX.utils.aoa_to_sheet(state.lastMatrixAOA);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'BiCubo_Matriz');
  const ext = bookType === 'xls' ? 'xls' : 'xlsx';
  XLSX.writeFile(wb, `relatorio_bi_cubo.${ext}`, { bookType: bookType === 'xls' ? 'biff8' : 'xlsx' });
}

document.getElementById('btnExportXlsx').addEventListener('click', () => exportMatrix('xlsx'));
document.getElementById('btnExportXls').addEventListener('click', () => exportMatrix('xls'));

document.getElementById('btnExportPdf').addEventListener('click', () => {
  if (!state.lastMatrixAOA) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape' });

  doc.setFontSize(16);
  doc.setTextColor(30, 58, 138);
  doc.text('RELATÓRIO EXECUTIVO BI(CUBO) ENTERPRISE', 14, 16);

  doc.setFontSize(9);
  doc.setTextColor(100, 116, 139);
  doc.text(`Gerado em: ${new Date().toLocaleString('pt-BR')}  |  Ambiente Corporativo OLAP`, 14, 22);

  const [head, ...body] = state.lastMatrixAOA;

  doc.autoTable({
    head: [head],
    body: body,
    startY: 28,
    styles: { fontSize: 7, cellPadding: 2, font: 'helvetica' },
    headStyles: { fillColor: [30, 58, 138], textColor: 255, fontStyle: 'bold' },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    theme: 'grid'
  });

  doc.save('relatorio_bi_cubo_executivo.pdf');
});

document.getElementById('btnExportChartPng').addEventListener('click', () => {
  const canvas = document.getElementById('chartCanvas');
  if (!canvas) return;
  const a = document.createElement('a');
  a.href = canvas.toDataURL('image/png');
  a.download = 'grafico_bi_cubo.png';
  a.click();
});

document.getElementById('btnExportChartPdf').addEventListener('click', () => {
  const canvas = document.getElementById('chartCanvas');
  if (!canvas) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape' });
  doc.setFontSize(16);
  doc.setTextColor(30, 58, 138);
  doc.text('GRÁFICO ANALÍTICO BI(CUBO)', 14, 16);

  const img = canvas.toDataURL('image/png');
  const pageWidth = doc.internal.pageSize.getWidth();
  const w = pageWidth - 28;
  const h = w * (canvas.height / canvas.width);
  doc.addImage(img, 'PNG', 14, 24, w, h);
  doc.save('grafico_bi_cubo.pdf');
});

// REFRESH ALL VIEWS
function refreshAll() {
  renderFieldPool();
  renderAllZones();
  renderMatrix();
  renderChart();
  updateExportButtons();
}

// SQL CONSOLE CONTROLLER
const btnOpenSqlModal = document.getElementById('btnOpenSqlModal');
const closeSqlBtn = document.getElementById('closeSqlBtn');
const btnCancelSql = document.getElementById('btnCancelSql');
const sqlModal = document.getElementById('sqlModal');
const btnRunSql = document.getElementById('btnRunSql');
const sqlPresets = document.getElementById('sqlPresets');
const sqlQueryText = document.getElementById('sqlQueryText');

if (btnOpenSqlModal) {
  btnOpenSqlModal.addEventListener('click', () => sqlModal.classList.remove('hidden'));
}
if (closeSqlBtn) {
  closeSqlBtn.addEventListener('click', () => sqlModal.classList.add('hidden'));
}
if (btnCancelSql) {
  btnCancelSql.addEventListener('click', () => sqlModal.classList.add('hidden'));
}

if (sqlPresets) {
  sqlPresets.addEventListener('change', () => {
    if (sqlPresets.value === 'all') {
      sqlQueryText.value = 'SELECT Mês, Família, Descrição_Item, Segmento, Região, Vendedor, Faturamento, Quantidade, Margem FROM Vendas_Empresariais;';
    } else if (sqlPresets.value === 'summary') {
      sqlQueryText.value = 'SELECT Mês, Família, SUM(Faturamento) AS total_fat FROM Vendas_Empresariais GROUP BY Mês, Família;';
    } else if (sqlPresets.value === 'top_items') {
      sqlQueryText.value = 'SELECT Descrição_Item, Segmento, SUM(Faturamento) AS total_fat FROM Vendas_Empresariais GROUP BY Descrição_Item, Segmento;';
    }
  });
}

if (btnRunSql) {
  btnRunSql.addEventListener('click', async () => {
    const statusEl = document.getElementById('sqlStatusMessage');
    const sql = (sqlQueryText?.value || '').trim();
    if (!sql) {
      alert('Digite uma consulta SQL.');
      return;
    }

    if (!state.data.length) {
      await loadDataset(generateDemoDataset(), 'Consulta SQL (Demo)');
    }

    if (statusEl) {
      statusEl.className = 'status-msg info';
      statusEl.textContent = 'Executando SELECT na API...';
    }

    try {
      const data = await api('/sql', {
        method: 'POST',
        body: {
          sql,
          datasetId: state.currentDatasetId,
          rows: state.currentDatasetId ? undefined : state.data,
        },
      });

      const resultRows = data.result?.rows || [];
      if (!resultRows.length) {
        if (statusEl) {
          statusEl.className = 'status-msg info';
          statusEl.textContent = 'Consulta OK — 0 linhas retornadas.';
        }
        alert('Consulta executada, mas não retornou linhas.');
        return;
      }

      await loadDataset(resultRows, `SQL · ${resultRows.length} linhas`, { persist: true });
      sqlModal.classList.add('hidden');
      if (statusEl) {
        statusEl.className = 'status-msg success';
        statusEl.textContent = `OK · ${resultRows.length} linhas · ${data.result.executedSql}`;
      }
    } catch (err) {
      if (statusEl) {
        statusEl.className = 'status-msg danger';
        statusEl.textContent = err.message || 'Falha SQL';
      }
      alert(err.message || 'Falha ao executar SQL');
    }
  });
}

// PERSISTENT SAVED CUBES (API + fallback local)
const LS_CUBES_KEY = 'bi_cubo_enterprise_saved_cubes_v2';

function loadLocalCubes() {
  try {
    const raw = localStorage.getItem(LS_CUBES_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function persistLocalCubes(list) {
  try {
    localStorage.setItem(LS_CUBES_KEY, JSON.stringify(list));
  } catch {}
}

async function fetchCloudCubes() {
  if (!getSession()) return null;
  try {
    const data = await api('/cubes');
    return data.cubes || [];
  } catch {
    return null;
  }
}

async function renderSavedCubes() {
  const host = document.getElementById('savedCubesList');
  if (!host) return;

  host.innerHTML = '<div class="empty-cubes-hint" style="font-size:0.7rem; color:var(--color-text-muted); padding:0.4rem 0;">Carregando...</div>';

  let list = await fetchCloudCubes();
  let source = 'nuvem';
  if (!list) {
    list = loadLocalCubes();
    source = 'local';
  }

  host.innerHTML = '';

  if (!list.length) {
    host.innerHTML = `<div class="empty-cubes-hint" style="font-size:0.7rem; color:var(--color-text-muted); padding:0.4rem 0;">Nenhum cubo salvo (${source}).</div>`;
    return;
  }

  list.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'cube-item-row';

    const titleSpan = document.createElement('span');
    titleSpan.textContent = item.name;
    titleSpan.title = `${item.name} (${new Date(item.savedAt).toLocaleDateString('pt-BR')}) · ${source}`;
    row.appendChild(titleSpan);

    const btnGroup = document.createElement('div');
    btnGroup.style.display = 'flex';
    btnGroup.style.gap = '0.2rem';

    const openBtn = document.createElement('button');
    openBtn.className = 'btn btn-xs btn-primary';
    openBtn.textContent = 'Abrir';
    openBtn.title = 'Carregar esta visão de cubo';
    openBtn.addEventListener('click', () => applySavedCube(item));
    btnGroup.appendChild(openBtn);

    const delBtn = document.createElement('button');
    delBtn.className = 'btn btn-xs btn-outline';
    delBtn.style.color = 'var(--color-danger)';
    delBtn.style.borderColor = 'var(--border-color)';
    delBtn.textContent = '✕';
    delBtn.title = 'Excluir cubo salvo';
    delBtn.addEventListener('click', async () => {
      if (source === 'nuvem' && item.id) {
        try {
          await api(`/cubes?id=${encodeURIComponent(item.id)}`, { method: 'DELETE' });
        } catch (err) {
          alert(err.message || 'Falha ao excluir');
          return;
        }
      } else {
        const cubes = loadLocalCubes().filter((c) => c.id !== item.id && c.name !== item.name);
        persistLocalCubes(cubes);
      }
      renderSavedCubes();
    });
    btnGroup.appendChild(delBtn);

    row.appendChild(btnGroup);
    host.appendChild(row);
  });
}

async function renderCloudDatasets() {
  const host = document.getElementById('cloudDatasetsList');
  if (!host || !getSession()) {
    if (host) host.innerHTML = '';
    return;
  }

  try {
    const data = await api('/datasets');
    const list = data.datasets || [];
    host.innerHTML = '';
    if (!list.length) {
      host.innerHTML = '<div class="empty-cubes-hint" style="font-size:0.7rem; color:var(--color-text-muted);">Nenhum dataset na nuvem.</div>';
      return;
    }
    list.forEach((ds) => {
      const row = document.createElement('div');
      row.className = 'cube-item-row';
      const title = document.createElement('span');
      title.textContent = `${ds.name} (${ds.rowCount})`;
      title.title = 'Carregar dataset da nuvem';
      row.appendChild(title);

      const openBtn = document.createElement('button');
      openBtn.className = 'btn btn-xs btn-primary';
      openBtn.textContent = 'Abrir';
      openBtn.addEventListener('click', async () => {
        try {
          const full = await api(`/dataset?id=${encodeURIComponent(ds.id)}`);
          await loadDataset(full.dataset.rows, full.dataset.name, {
            persist: false,
            datasetId: full.dataset.id,
          });
        } catch (err) {
          alert(err.message || 'Falha ao abrir dataset');
        }
      });
      row.appendChild(openBtn);
      host.appendChild(row);
    });
  } catch {
    host.innerHTML = '';
  }
}

function applySavedCube(item) {
  if (!state.fields.length) {
    alert('Por favor, carregue uma base de dados antes de abrir um cubo salvo.');
    return;
  }

  const validField = (f) => state.fields.includes(f);

  state.pivot.rows = (item.pivot.rows || []).filter(validField);
  state.pivot.cols = (item.pivot.cols || []).filter(validField);
  state.pivot.filters = (item.pivot.filters || []).filter(validField);
  state.pivot.values = (item.pivot.values || []).filter(v => validField(v.field));

  if (item.fieldAliases) {
    state.fieldAliases = { ...item.fieldAliases };
  }
  if (item.fieldSortDirs) {
    state.fieldSortDirs = { ...item.fieldSortDirs };
  }

  if (item.filterSelected) {
    Object.keys(item.filterSelected).forEach((f) => {
      if (validField(f)) {
        state.filterSelected[f] = new Set(item.filterSelected[f]);
      }
    });
  }

  if (item.chart) {
    state.chart = {
      x: validField(item.chart.x) ? item.chart.x : null,
      series: validField(item.chart.series) ? item.chart.series : null,
      value: validField(item.chart.value) ? item.chart.value : null,
      agg: item.chart.agg || 'sum'
    };
    if (item.chartType) {
      const chartTypeSelect = document.getElementById('chartType');
      if (chartTypeSelect) chartTypeSelect.value = item.chartType;
    }
  }

  refreshAll();
  alert(`Cubo "${item.name}" carregado com sucesso!`);
}

const btnSaveCube = document.getElementById('btnSaveCube');
if (btnSaveCube) {
  btnSaveCube.addEventListener('click', async () => {
    if (!state.data.length) {
      alert('Carregue dados antes de salvar o cubo.');
      return;
    }

    const defaultName = `Cubo - ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`;
    const cubeName = prompt('Digite um nome para este cubo (visão analítica):', defaultName);
    if (!cubeName || !cubeName.trim()) return;

    const filterSelectedObj = {};
    Object.keys(state.filterSelected).forEach((f) => {
      if (state.filterSelected[f]) {
        filterSelectedObj[f] = Array.from(state.filterSelected[f]);
      }
    });

    const cubeRecord = {
      name: cubeName.trim(),
      datasetId: state.currentDatasetId,
      pivot: {
        rows: state.pivot.rows,
        cols: state.pivot.cols,
        filters: state.pivot.filters,
        values: state.pivot.values
      },
      fieldAliases: state.fieldAliases,
      fieldSortDirs: state.fieldSortDirs,
      filterSelected: filterSelectedObj,
      chart: state.chart,
      chartType: document.getElementById('chartType') ? document.getElementById('chartType').value : 'bar',
    };

    try {
      if (getSession()) {
        await api('/cubes', { method: 'POST', body: cubeRecord });
        alert(`Cubo "${cubeName.trim()}" salvo na nuvem!`);
      } else {
        const local = loadLocalCubes();
        local.push({ ...cubeRecord, id: 'local_' + Date.now(), savedAt: Date.now() });
        persistLocalCubes(local);
        alert(`Cubo "${cubeName.trim()}" salvo localmente neste navegador.`);
      }
      renderSavedCubes();
    } catch (err) {
      alert(err.message || 'Falha ao salvar cubo');
    }
  });
}

// INIT AUTH & APP
document.addEventListener('DOMContentLoaded', () => {
  initAuth(async () => {
    if (!state.data.length) {
      await loadDataset(generateDemoDataset(), 'Vendas Empresariais (Demo)');
    }
    renderSavedCubes();
    renderCloudDatasets();
  });
});
