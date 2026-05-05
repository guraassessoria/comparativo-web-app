/* global XLSX */
(() => {
  const state = {
    workbook: null,
    workbookName: '',
    rowsBySheet: new Map(),
    rawRows: [],
    headerRowIndex: -1,
    columnDefs: [],
    defaultColumns: null,
    baseItems: [],
    analysis: null,
  };

  const els = {
    fileInput: document.getElementById('fileInput'),
    dropZone: document.getElementById('dropZone'),
    sheetSelect: document.getElementById('sheetSelect'),
    periodModeSelect: document.getElementById('periodModeSelect'),
    dateColumnSelect: document.getElementById('dateColumnSelect'),
    historyColumnSelect: document.getElementById('historyColumnSelect'),
    debitColumnSelect: document.getElementById('debitColumnSelect'),
    creditColumnSelect: document.getElementById('creditColumnSelect'),
    valueColumnSelect: document.getElementById('valueColumnSelect'),
    filterColumnsSelect: document.getElementById('filterColumnsSelect'),
    removeDuplicatesToggle: document.getElementById('removeDuplicatesToggle'),
    analyzeBtn: document.getElementById('analyzeBtn'),
    exportBtn: document.getElementById('exportBtn'),
    resetBtn: document.getElementById('resetBtn'),
    statusBox: document.getElementById('statusBox'),
    summarySection: document.getElementById('summarySection'),
    previewSection: document.getElementById('previewSection'),
    previewTable: document.getElementById('previewTable'),
    searchInput: document.getElementById('searchInput'),
    filterControls: document.getElementById('filterControls'),
    rulesDialog: document.getElementById('rulesDialog'),
    sampleRulesBtn: document.getElementById('sampleRulesBtn'),
    metricComparison: document.getElementById('metricComparison'),
    metricPeriodMode: document.getElementById('metricPeriodMode'),
    metricUsed: document.getElementById('metricUsed'),
    metricDuplicates: document.getElementById('metricDuplicates'),
    metricNames: document.getElementById('metricNames'),
    metricVariation: document.getElementById('metricVariation'),
  };

  const DEFAULT_COLUMNS = {
    data: 0,
    categoria: 1,
    centro: 2,
    historico: 3,
    debito: 7,
    credito: 8,
  };

  const PERIOD_LABELS = {
    monthly: 'Mensal',
    quarterly: 'Trimestral',
    semiannual: 'Semestral',
    annual: 'Anual',
  };

  const MONTH_NAMES = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
  ];

  const SERVICE_MARKERS = [
    ' HIST. ',
    ' HIST ',
    ' HISTORICO ',
    ' HI ST. ',
    ' HIST SERVICOS ',
    ' HIST. SERVICOS ',
  ];

  function normalizeText(value) {
    return String(value ?? '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[“”]/g, '"')
      .replace(/[’]/g, "'")
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
  }

  function normalizeCompact(value) {
    return normalizeText(value).replace(/[^A-Z0-9]+/g, ' ').trim();
  }

  function cleanName(value) {
    let name = normalizeText(value);
    name = name.replace(/^[-–—:\s]+/, '');
    name = name.replace(/\s+/g, ' ').trim();

    for (const marker of SERVICE_MARKERS) {
      const idx = name.indexOf(marker);
      if (idx > 0) name = name.slice(0, idx).trim();
    }

    name = name
      .replace(/\bHIST\.?\s*SERV.*$/i, '')
      .replace(/\bHI\s*ST\.?\s*SERV.*$/i, '')
      .replace(/\bSERVICOS?\s+DE\s+(ASSESSORIA|CONSULTORIA|AD\s*VOGADOS|ADVOGADOS|IN\s*FORMATICA|INFORMATICA).*$/i, '')
      .replace(/\bOUTROS\s+SERVICOS\s+PROFISSIONAIS.*$/i, '')
      .replace(/\s+HI\s*ST\.?$/i, '')
      .replace(/\s+HIST\.?$/i, '')
      .replace(/\s+20\d{4}$/i, '')
      .replace(/\s+\d{4,6}$/i, '')
      .replace(/[.;,]+$/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    return name || 'NAO IDENTIFICADO';
  }

  function colLetter(index) {
    let n = index + 1;
    let result = '';
    while (n > 0) {
      const mod = (n - 1) % 26;
      result = String.fromCharCode(65 + mod) + result;
      n = Math.floor((n - mod) / 26);
    }
    return result;
  }

  function parseNumber(value) {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'number' && Number.isFinite(value)) return value;

    let text = String(value).trim();
    if (!text) return 0;

    const isCreditLike = /\bC\b$/i.test(text);
    text = text.replace(/[R$\s]/g, '').replace(/[DC]$/i, '');

    if (text.includes(',') && text.includes('.')) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else if (text.includes(',')) {
      text = text.replace(',', '.');
    }

    const parsed = Number(text.replace(/[^0-9.-]/g, ''));
    if (!Number.isFinite(parsed)) return 0;
    return isCreditLike ? -Math.abs(parsed) : parsed;
  }

  function excelSerialToDate(serial) {
    if (!Number.isFinite(serial)) return null;
    const parsed = XLSX.SSF.parse_date_code(serial);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === 'number') return excelSerialToDate(value);
    const text = String(value ?? '').trim();
    if (!text) return null;

    let match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));

    match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:\s+.*)?$/);
    if (match) {
      let year = Number(match[3]);
      if (year < 100) year += 2000;
      return new Date(year, Number(match[2]) - 1, Number(match[1]));
    }

    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatDate(value) {
    const date = parseDate(value);
    if (!date) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function formatDateBR(value) {
    const date = parseDate(value);
    if (!date) return '';
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    return `${d}/${m}/${date.getFullYear()}`;
  }

  function formatMoney(value) {
    return Number(value || 0).toLocaleString('pt-BR', {
      style: 'currency',
      currency: 'BRL',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  function formatNumber(value) {
    return Number(value || 0).toLocaleString('pt-BR');
  }

  function formatPercent(value) {
    if (value === null || value === undefined || Number.isNaN(value)) return '-';
    return Number(value).toLocaleString('pt-BR', {
      style: 'percent',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  function round2(value) {
    return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
  }

  function setStatus(message, type = 'muted') {
    els.statusBox.textContent = message;
    els.statusBox.className = `status-box ${type}`;
  }

  function findHeaderRow(rows) {
    const maxRows = Math.min(rows.length, 30);
    for (let r = 0; r < maxRows; r += 1) {
      const normalized = (rows[r] || []).map(normalizeCompact);
      const hasHistorico = normalized.some((cell) => cell.includes('HISTORICO') || cell === 'HIST');
      const hasData = normalized.some((cell) => cell === 'DATA' || cell.includes('DT '));
      const hasDebit = normalized.some((cell) => cell.includes('DEBITO'));
      const hasCredit = normalized.some((cell) => cell.includes('CREDITO'));
      if (hasHistorico && (hasData || hasDebit || hasCredit)) return r;
    }
    return -1;
  }

  function getMaxColumnCount(rows) {
    return rows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
  }

  function sampleColumnValues(rows, index, headerRowIndex) {
    const start = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
    const samples = [];
    for (let i = start; i < rows.length && samples.length < 3; i += 1) {
      const value = rows[i]?.[index];
      const text = String(value ?? '').trim();
      if (text) samples.push(text.length > 28 ? `${text.slice(0, 28)}...` : text);
    }
    return samples.join(' | ');
  }

  function buildColumnDefs(rows, headerRowIndex) {
    const maxCols = getMaxColumnCount(rows);
    const headerRow = headerRowIndex >= 0 ? rows[headerRowIndex] || [] : [];
    const usedLabels = new Map();

    return Array.from({ length: maxCols }, (_, index) => {
      const letter = colLetter(index);
      const rawHeader = String(headerRow[index] ?? '').trim();
      const sample = sampleColumnValues(rows, index, headerRowIndex);
      let label = rawHeader || `Coluna ${letter}`;
      label = label.replace(/\s+/g, ' ').trim();
      const count = usedLabels.get(label) || 0;
      usedLabels.set(label, count + 1);
      const uniqueLabel = count ? `${label} (${letter})` : label;
      const optionText = rawHeader
        ? `${letter} — ${uniqueLabel}${sample ? ` | ${sample}` : ''}`
        : `${letter} — ${sample || 'sem amostra'}`;
      return { index, letter, label: uniqueLabel, optionText };
    });
  }

  function detectDefaultColumns(rows, headerRowIndex, columnDefs) {
    const headerRow = headerRowIndex >= 0 ? rows[headerRowIndex] || [] : [];
    const normalized = columnDefs.map((def) => normalizeCompact(headerRow[def.index] ?? def.label));
    const findIndex = (predicates) => {
      const found = normalized.findIndex((cell) => predicates.some((fn) => fn(cell)));
      return found >= 0 ? found : -1;
    };

    return {
      data: findIndex([(cell) => cell === 'DATA', (cell) => cell.includes('DATA'), (cell) => cell.includes('DT')]),
      historico: findIndex([(cell) => cell.includes('HISTORICO'), (cell) => cell === 'HIST']),
      debito: findIndex([(cell) => cell.includes('DEBITO')]),
      credito: findIndex([(cell) => cell.includes('CREDITO')]),
      valor: findIndex([(cell) => cell === 'VALOR', (cell) => cell.includes('VALOR LIQUIDO'), (cell) => cell.includes('VLR LIQUIDO')]),
      categoria: findIndex([
        (cell) => cell.includes('CATEGORIA'),
        (cell) => cell.includes('TIPO'),
        (cell) => cell.includes('CONTA'),
        (cell) => cell.includes('SERVICO'),
      ]),
      centro: findIndex([
        (cell) => cell.includes('CENTRO'),
        (cell) => cell.includes('AREA'),
        (cell) => cell.includes('DEPARTAMENTO'),
        (cell) => cell.includes('CUSTO'),
      ]),
    };
  }

  function addColumnOptions(select, columnDefs, includeBlank = false, blankLabel = 'Não usar') {
    select.innerHTML = '';
    if (includeBlank) {
      const blank = document.createElement('option');
      blank.value = '';
      blank.textContent = blankLabel;
      select.appendChild(blank);
    }
    columnDefs.forEach((def) => {
      const option = document.createElement('option');
      option.value = String(def.index);
      option.textContent = def.optionText;
      select.appendChild(option);
    });
  }

  function setSelectValue(select, index, fallbackIndex = null) {
    const desired = Number.isInteger(index) && index >= 0 ? index : fallbackIndex;
    if (desired !== null && desired !== undefined && desired >= 0) select.value = String(desired);
  }

  function setMultiSelectValues(select, indexes) {
    const values = new Set(indexes.filter((idx) => Number.isInteger(idx) && idx >= 0).map(String));
    Array.from(select.options).forEach((option) => {
      option.selected = values.has(option.value);
    });
  }

  function enableMappingControls(enabled) {
    [
      els.periodModeSelect,
      els.dateColumnSelect,
      els.historyColumnSelect,
      els.debitColumnSelect,
      els.creditColumnSelect,
      els.valueColumnSelect,
      els.filterColumnsSelect,
    ].forEach((el) => { el.disabled = !enabled; });
  }

  function loadSheet(sheetName) {
    const rows = state.rowsBySheet.get(sheetName) || [];
    state.rawRows = rows;
    state.headerRowIndex = findHeaderRow(rows);
    state.columnDefs = buildColumnDefs(rows, state.headerRowIndex);
    state.baseItems = [];
    state.analysis = null;

    const defaults = detectDefaultColumns(rows, state.headerRowIndex, state.columnDefs);
    const fallback = {
      data: defaults.data >= 0 ? defaults.data : DEFAULT_COLUMNS.data,
      historico: defaults.historico >= 0 ? defaults.historico : DEFAULT_COLUMNS.historico,
      debito: defaults.debito >= 0 ? defaults.debito : DEFAULT_COLUMNS.debito,
      credito: defaults.credito >= 0 ? defaults.credito : DEFAULT_COLUMNS.credito,
      categoria: defaults.categoria >= 0 ? defaults.categoria : DEFAULT_COLUMNS.categoria,
      centro: defaults.centro >= 0 ? defaults.centro : DEFAULT_COLUMNS.centro,
    };
    state.defaultColumns = fallback;

    addColumnOptions(els.dateColumnSelect, state.columnDefs);
    addColumnOptions(els.historyColumnSelect, state.columnDefs);
    addColumnOptions(els.debitColumnSelect, state.columnDefs, true);
    addColumnOptions(els.creditColumnSelect, state.columnDefs, true);
    addColumnOptions(els.valueColumnSelect, state.columnDefs, true, 'Não usar: calcular Débito - Crédito');
    addColumnOptions(els.filterColumnsSelect, state.columnDefs);

    setSelectValue(els.dateColumnSelect, defaults.data, fallback.data);
    setSelectValue(els.historyColumnSelect, defaults.historico, fallback.historico);
    setSelectValue(els.debitColumnSelect, defaults.debito, fallback.debito);
    setSelectValue(els.creditColumnSelect, defaults.credito, fallback.credito);
    setSelectValue(els.valueColumnSelect, defaults.valor, null);
    setMultiSelectValues(els.filterColumnsSelect, [fallback.categoria, fallback.centro]);

    enableMappingControls(state.columnDefs.length > 0);
    els.analyzeBtn.disabled = state.columnDefs.length === 0;
    els.exportBtn.disabled = true;
    els.summarySection.classList.add('hidden');
    els.previewSection.classList.add('hidden');
    els.filterControls.classList.add('hidden');
    els.filterControls.innerHTML = '';

    const headerInfo = state.headerRowIndex >= 0
      ? `Cabeçalho detectado na linha ${state.headerRowIndex + 1}. Confira as colunas antes de analisar.`
      : 'Sem cabeçalho detectado; as opções foram montadas com letras e amostras das colunas.';
    setStatus(`${formatNumber(rows.length)} linhas lidas da aba selecionada. ${headerInfo}`, 'success');
  }

  function getSelectedFilterDefs() {
    const selected = Array.from(els.filterColumnsSelect.selectedOptions).map((option) => Number(option.value));
    return selected
      .filter((idx) => Number.isInteger(idx) && idx >= 0)
      .map((idx) => state.columnDefs.find((def) => def.index === idx))
      .filter(Boolean);
  }

  function getSelectedMapping() {
    const data = Number(els.dateColumnSelect.value);
    const historico = Number(els.historyColumnSelect.value);
    const debito = els.debitColumnSelect.value === '' ? null : Number(els.debitColumnSelect.value);
    const credito = els.creditColumnSelect.value === '' ? null : Number(els.creditColumnSelect.value);
    const valor = els.valueColumnSelect.value === '' ? null : Number(els.valueColumnSelect.value);
    const filterDefs = getSelectedFilterDefs();
    const categoryIndex = state.defaultColumns?.categoria ?? null;

    if (!Number.isInteger(data) || data < 0) throw new Error('Selecione uma coluna de data válida.');
    if (!Number.isInteger(historico) || historico < 0) throw new Error('Selecione uma coluna de histórico válida.');
    if (valor === null && debito === null && credito === null) {
      throw new Error('Selecione uma coluna de valor pronto ou pelo menos uma coluna de débito/crédito.');
    }

    return { data, historico, debito, credito, valor, filterDefs, categoryIndex };
  }

  function normalizeCategory(category) {
    const text = normalizeCompact(category);
    if (/\bAUTONOMOS?\b/.test(text)) return 'AUTONOMO';
    if (/\bINSS\b.*\bS\b.*\bPF\b/.test(text)) return 'AUTONOMO';
    if (text === 'INSS S PF') return 'AUTONOMO';
    return normalizeText(category) || 'SEM CATEGORIA';
  }

  function inferCategoryFromValues(values) {
    const joined = values.map(normalizeText).join(' | ');
    return normalizeCategory(joined);
  }

  function extractNF(history) {
    const text = normalizeText(history);
    const match = text.match(/\bNF\.?\s*(?:N[ºO]?\.?\s*)?0*(\d{1,12})\b/i);
    return match ? String(Number(match[1])) : '';
  }

  function extractNameAfterNF(history) {
    const original = normalizeText(history);
    const match = original.match(/\bNF\.?\s*(?:N[ºO]?\.?\s*)?0*\d{1,12}\s*[-–—:]?\s*(.+)$/i);
    if (!match) return '';
    return cleanName(match[1]);
  }

  function extractNameFromRefNF(history) {
    const text = normalizeText(history);
    const isRefNF = /\bREF\.?\s+NF\.?\b/i.test(text) || /^NF\.?\s*0*\d+/i.test(text);
    if (!isRefNF) return '';

    const afterNF = extractNameAfterNF(text);
    if (!afterNF || afterNF === 'NAO IDENTIFICADO') return '';
    return afterNF;
  }

  function extractAutonomoName(history) {
    const text = normalizeText(history);

    if (/REVER.*IRRF|RECLASSIFICACAO.*IRRF|RECLASSIFICAÇÃO.*IRRF/.test(text)) {
      return { name: 'AUTONOMOS - AJUSTES IRRF', method: 'AJUSTE_AUTONOMO' };
    }

    const provMatch = text.match(/\bPROV\.?\s*(?:FOL\.?|FOLHA)?\s*(?:A\s+MAIOR\s+)?(.+?)(?:\s+\d{2,6}|\s+20\d{2}|$)/i);
    if (provMatch && provMatch[1]) {
      const name = cleanName(provMatch[1]);
      if (name && name !== 'NAO IDENTIFICADO') return { name, method: 'PROVISAO_AUTONOMO' };
    }

    return { name: 'AUTONOMOS', method: 'AUTONOMO_GENERICO' };
  }

  function buildNFMaps(items) {
    const byCategoryAndNF = new Map();
    const byNF = new Map();

    function setBest(map, key, name) {
      if (!key || !name || name === 'NAO IDENTIFICADO') return;
      const current = map.get(key);
      if (!current || name.length > current.length) map.set(key, name);
    }

    items.forEach((item) => {
      const nf = extractNF(item.historico);
      if (!nf) return;
      const name = extractNameFromRefNF(item.historico);
      if (!name) return;
      setBest(byCategoryAndNF, `${item.categoriaTratada}|${nf}`, name);
      setBest(byNF, nf, name);
    });

    return { byCategoryAndNF, byNF };
  }

  function getLookupName(nfMaps, category, nf) {
    if (!nf) return '';
    return nfMaps.byCategoryAndNF.get(`${category}|${nf}`) || nfMaps.byNF.get(nf) || '';
  }

  function inferNameAndMethod(item, nfMaps) {
    const history = item.historico;
    const text = normalizeText(history);
    const nf = extractNF(history);

    if (item.categoriaTratada === 'AUTONOMO') {
      const aut = extractAutonomoName(history);
      return { name: aut.name, nf, method: aut.method };
    }

    const shouldLookup = /\b(RETIDO|IRRF|ISS|INSS|PCC|REVERS|REVERSAO|REVERSÃO|VLR\s+REF)\b/i.test(text);
    if (shouldLookup) {
      const lookupName = getLookupName(nfMaps, item.categoriaTratada, nf);
      if (lookupName) return { name: lookupName, nf, method: 'RETIDO_LOOKUP' };
    }

    const refName = extractNameFromRefNF(history);
    if (refName) return { name: refName, nf, method: 'REF_NF' };

    if (nf) {
      const lookupName = getLookupName(nfMaps, item.categoriaTratada, nf);
      if (lookupName) return { name: lookupName, nf, method: 'AJUSTE_LOOKUP' };

      const afterNF = extractNameAfterNF(history);
      if (afterNF && afterNF !== 'NAO IDENTIFICADO') return { name: afterNF, nf, method: 'AJUSTE_TEXTO' };
    }

    return { name: cleanName(history), nf, method: 'HISTORICO_FALLBACK' };
  }

  function buildFilterValues(row, filterDefs) {
    const out = {};
    filterDefs.forEach((def) => {
      out[def.label] = normalizeText(row[def.index]) || '';
    });
    return out;
  }

  function periodMeta(date, mode) {
    if (!date) return { key: 'SEM_PERIODO', label: 'Sem período', order: 999 };
    const month = date.getMonth() + 1;
    if (mode === 'monthly') {
      return {
        key: `M${String(month).padStart(2, '0')}`,
        label: `${String(month).padStart(2, '0')} - ${MONTH_NAMES[month - 1]}`,
        order: month,
      };
    }
    if (mode === 'quarterly') {
      const q = Math.ceil(month / 3);
      return { key: `T${q}`, label: `${q}º Trimestre`, order: q };
    }
    if (mode === 'semiannual') {
      const s = month <= 6 ? 1 : 2;
      return { key: `S${s}`, label: `${s}º Semestre`, order: s };
    }
    return { key: 'ANUAL', label: 'Anual', order: 1 };
  }

  function rowToItem(row, sourceIndex, mapping, periodMode) {
    const rawDate = row[mapping.data];
    const date = parseDate(rawDate);
    const filterValues = buildFilterValues(row, mapping.filterDefs);
    const categoryCandidates = Object.values(filterValues || []);
    if (Number.isInteger(mapping.categoryIndex) && mapping.categoryIndex >= 0) {
      categoryCandidates.unshift(row[mapping.categoryIndex]);
    }
    const categoriaTratada = inferCategoryFromValues(categoryCandidates);
    const historico = normalizeText(row[mapping.historico]);
    const debito = mapping.debito === null ? 0 : parseNumber(row[mapping.debito]);
    const credito = mapping.credito === null ? 0 : parseNumber(row[mapping.credito]);
    const valorLiquido = mapping.valor === null ? debito - credito : parseNumber(row[mapping.valor]);
    const period = periodMeta(date, periodMode);

    return {
      linhaOriginal: sourceIndex + 1,
      data: date,
      dataISO: formatDate(date),
      ano: date ? date.getFullYear() : null,
      mes: date ? date.getMonth() + 1 : null,
      periodoChave: period.key,
      periodo: period.label,
      periodoOrdem: period.order,
      categoriaTratada,
      filterValues,
      historico,
      debito,
      credito,
      valorLiquido,
      nf: '',
      nomeExtraido: '',
      metodoExtracao: '',
      excluidaDuplicidade: false,
    };
  }

  function buildBaseItems(rows, mapping, periodMode) {
    const start = state.headerRowIndex >= 0 ? state.headerRowIndex + 1 : 0;
    const usableRows = rows.slice(start);
    const preliminary = [];

    usableRows.forEach((row, offset) => {
      const sourceIndex = start + offset;
      if (!row || row.every((cell) => cell === null || cell === undefined || String(cell).trim() === '')) return;
      const item = rowToItem(row, sourceIndex, mapping, periodMode);
      if (!item.historico && !item.data) return;
      preliminary.push(item);
    });

    const nfMaps = buildNFMaps(preliminary);
    preliminary.forEach((item) => {
      const extracted = inferNameAndMethod(item, nfMaps);
      item.nf = extracted.nf;
      item.nomeExtraido = extracted.name;
      item.metodoExtracao = extracted.method;
    });

    return preliminary;
  }

  function getAvailableYears(items) {
    return [...new Set(items.map((item) => item.ano).filter(Boolean))].sort((a, b) => b - a);
  }

  function rowSignature(item, filterDefs) {
    const filterPart = filterDefs.map((def) => normalizeText(item.filterValues[def.label])).join('¦');
    return [
      item.dataISO,
      filterPart,
      normalizeText(item.historico),
      Number(item.debito || 0).toFixed(2),
      Number(item.credito || 0).toFixed(2),
      Number(item.valorLiquido || 0).toFixed(2),
    ].join('¦');
  }

  function findDuplicateBlocks(items, filterDefs, minRun = 20) {
    const signatures = items.map((item) => rowSignature(item, filterDefs));
    const positions = new Map();
    signatures.forEach((sig, idx) => {
      if (!positions.has(sig)) positions.set(sig, []);
      positions.get(sig).push(idx);
    });

    const removed = new Set();
    const blocks = [];

    for (let i = 0; i < signatures.length; i += 1) {
      if (removed.has(i)) continue;
      const candidates = positions.get(signatures[i]) || [];
      for (const j of candidates) {
        if (j <= i || removed.has(j)) continue;
        let run = 0;
        while (
          i + run < signatures.length &&
          j + run < signatures.length &&
          !removed.has(j + run) &&
          signatures[i + run] === signatures[j + run]
        ) {
          run += 1;
        }
        if (run >= minRun) {
          for (let k = 0; k < run; k += 1) removed.add(j + k);
          blocks.push({
            Bloco_Original_Inicio: items[i].linhaOriginal,
            Bloco_Original_Fim: items[i + run - 1].linhaOriginal,
            Bloco_Duplicado_Inicio: items[j].linhaOriginal,
            Bloco_Duplicado_Fim: items[j + run - 1].linhaOriginal,
            Linhas_Removidas: run,
          });
          break;
        }
      }
    }

    return { removed, blocks };
  }

  function addDistinctValues(target, filterValues) {
    Object.entries(filterValues || {}).forEach(([key, value]) => {
      if (!target[key]) target[key] = new Set();
      if (value) target[key].add(value);
    });
  }

  function serializeFilterSets(filterSets) {
    const out = {};
    Object.entries(filterSets || {}).forEach(([key, set]) => {
      out[key] = [...set].sort().join(', ');
    });
    return out;
  }

  function buildStatus(previous, current, currentYear) {
    const variacao = current - previous;
    if (previous === 0 && current !== 0) return `Novo em ${currentYear}`;
    if (previous !== 0 && current === 0) return `Sem movimento em ${currentYear}`;
    if (variacao > 0) return 'Aumento';
    if (variacao < 0) return 'Redução';
    return 'Sem variação';
  }

  function sortComparisonRows(a, b, currentYear) {
    if ((a.Periodo_Ordem || 0) !== (b.Periodo_Ordem || 0)) return (a.Periodo_Ordem || 0) - (b.Periodo_Ordem || 0);
    return Math.abs(b[`Valor_${currentYear}`] || 0) - Math.abs(a[`Valor_${currentYear}`] || 0);
  }

  function groupByName(items, currentYear, previousYear, filterDefs) {
    const map = new Map();

    function ensure(item) {
      const key = `${item.periodoChave}|${item.nomeExtraido || 'NAO IDENTIFICADO'}`;
      if (!map.has(key)) {
        map.set(key, {
          Periodo_Ordem: item.periodoOrdem,
          Periodo: item.periodo,
          Nome_Extraido: item.nomeExtraido || 'NAO IDENTIFICADO',
          Filtros: {},
          Qtde_Anterior: 0,
          Qtde_Atual: 0,
          Valor_Anterior: 0,
          Valor_Atual: 0,
          Total_2_anos: 0,
        });
      }
      return map.get(key);
    }

    items.forEach((item) => {
      if (item.ano !== currentYear && item.ano !== previousYear) return;
      const row = ensure(item);
      addDistinctValues(row.Filtros, item.filterValues);
      row.Total_2_anos += item.valorLiquido;

      if (item.ano === previousYear) {
        row.Qtde_Anterior += 1;
        row.Valor_Anterior += item.valorLiquido;
      }
      if (item.ano === currentYear) {
        row.Qtde_Atual += 1;
        row.Valor_Atual += item.valorLiquido;
      }
    });

    return [...map.values()].map((row) => {
      const variacao = row.Valor_Atual - row.Valor_Anterior;
      const variacaoPct = row.Valor_Anterior !== 0 ? variacao / row.Valor_Anterior : null;
      const filterColumns = serializeFilterSets(row.Filtros);
      const output = {
        Periodo: row.Periodo,
        Nome_Extraido: row.Nome_Extraido,
      };
      filterDefs.forEach((def) => { output[def.label] = filterColumns[def.label] || ''; });
      return {
        Periodo_Ordem: row.Periodo_Ordem,
        ...output,
        [`Qtde_${previousYear}`]: row.Qtde_Anterior,
        [`Qtde_${currentYear}`]: row.Qtde_Atual,
        [`Valor_${previousYear}`]: round2(row.Valor_Anterior),
        [`Valor_${currentYear}`]: round2(row.Valor_Atual),
        'Variacao_R$': round2(variacao),
        'Variacao_%': variacaoPct === null ? null : variacaoPct,
        Total_2_anos: round2(row.Total_2_anos),
        Status: buildStatus(row.Valor_Anterior, row.Valor_Atual, currentYear),
      };
    }).sort((a, b) => sortComparisonRows(a, b, currentYear));
  }

  function groupByPeriod(items, currentYear, previousYear) {
    const map = new Map();
    items.forEach((item) => {
      if (item.ano !== currentYear && item.ano !== previousYear) return;
      const key = item.periodoChave;
      if (!map.has(key)) {
        map.set(key, {
          Periodo_Ordem: item.periodoOrdem,
          Periodo: item.periodo,
          [`Qtde_${previousYear}`]: 0,
          [`Qtde_${currentYear}`]: 0,
          [`Valor_${previousYear}`]: 0,
          [`Valor_${currentYear}`]: 0,
        });
      }
      const row = map.get(key);
      row[`Qtde_${item.ano}`] += 1;
      row[`Valor_${item.ano}`] += item.valorLiquido;
    });

    return [...map.values()].map((row) => {
      const variacao = row[`Valor_${currentYear}`] - row[`Valor_${previousYear}`];
      return {
        ...row,
        [`Valor_${previousYear}`]: round2(row[`Valor_${previousYear}`]),
        [`Valor_${currentYear}`]: round2(row[`Valor_${currentYear}`]),
        'Variacao_R$': round2(variacao),
        'Variacao_%': row[`Valor_${previousYear}`] !== 0 ? variacao / row[`Valor_${previousYear}`] : null,
      };
    }).sort((a, b) => (a.Periodo_Ordem || 0) - (b.Periodo_Ordem || 0));
  }

  function groupByCategory(items, currentYear, previousYear) {
    const map = new Map();
    items.forEach((item) => {
      if (item.ano !== currentYear && item.ano !== previousYear) return;
      const key = `${item.periodoChave}|${item.categoriaTratada || 'SEM CATEGORIA'}`;
      if (!map.has(key)) {
        map.set(key, {
          Periodo_Ordem: item.periodoOrdem,
          Periodo: item.periodo,
          Categoria_Tratada: item.categoriaTratada || 'SEM CATEGORIA',
          [`Qtde_${previousYear}`]: 0,
          [`Qtde_${currentYear}`]: 0,
          [`Valor_${previousYear}`]: 0,
          [`Valor_${currentYear}`]: 0,
        });
      }
      const row = map.get(key);
      row[`Qtde_${item.ano}`] += 1;
      row[`Valor_${item.ano}`] += item.valorLiquido;
    });

    return [...map.values()].map((row) => {
      const variacao = row[`Valor_${currentYear}`] - row[`Valor_${previousYear}`];
      return {
        ...row,
        [`Valor_${previousYear}`]: round2(row[`Valor_${previousYear}`]),
        [`Valor_${currentYear}`]: round2(row[`Valor_${currentYear}`]),
        'Variacao_R$': round2(variacao),
        'Variacao_%': row[`Valor_${previousYear}`] !== 0 ? variacao / row[`Valor_${previousYear}`] : null,
      };
    }).sort((a, b) => sortComparisonRows(a, b, currentYear));
  }

  function groupByFilters(items, currentYear, previousYear, filterDefs) {
    const map = new Map();
    filterDefs.forEach((def) => {
      items.forEach((item) => {
        if (item.ano !== currentYear && item.ano !== previousYear) return;
        const value = item.filterValues[def.label] || 'VAZIO';
        const key = `${def.label}|${value}|${item.periodoChave}`;
        if (!map.has(key)) {
          map.set(key, {
            Filtro: def.label,
            Valor_Filtro: value,
            Periodo_Ordem: item.periodoOrdem,
            Periodo: item.periodo,
            [`Qtde_${previousYear}`]: 0,
            [`Qtde_${currentYear}`]: 0,
            [`Valor_${previousYear}`]: 0,
            [`Valor_${currentYear}`]: 0,
          });
        }
        const row = map.get(key);
        row[`Qtde_${item.ano}`] += 1;
        row[`Valor_${item.ano}`] += item.valorLiquido;
      });
    });

    return [...map.values()].map((row) => {
      const variacao = row[`Valor_${currentYear}`] - row[`Valor_${previousYear}`];
      return {
        ...row,
        [`Valor_${previousYear}`]: round2(row[`Valor_${previousYear}`]),
        [`Valor_${currentYear}`]: round2(row[`Valor_${currentYear}`]),
        'Variacao_R$': round2(variacao),
        'Variacao_%': row[`Valor_${previousYear}`] !== 0 ? variacao / row[`Valor_${previousYear}`] : null,
      };
    }).sort((a, b) => {
      if (a.Filtro !== b.Filtro) return a.Filtro.localeCompare(b.Filtro);
      if ((a.Periodo_Ordem || 0) !== (b.Periodo_Ordem || 0)) return (a.Periodo_Ordem || 0) - (b.Periodo_Ordem || 0);
      return Math.abs(b[`Valor_${currentYear}`] || 0) - Math.abs(a[`Valor_${currentYear}`] || 0);
    });
  }

  function analyze() {
    const mapping = getSelectedMapping();
    const periodMode = els.periodModeSelect.value;
    const baseItems = buildBaseItems(state.rawRows, mapping, periodMode);
    if (!baseItems.length) throw new Error('Nenhuma linha válida encontrada na planilha. Confira as colunas selecionadas.');

    const years = getAvailableYears(baseItems);
    if (years.length < 2) {
      throw new Error('A coluna de data selecionada não contém pelo menos dois anos diferentes para comparar.');
    }

    const currentYear = years[0];
    const previousYear = years[1];
    const items = baseItems.map((item) => ({ ...item, excluidaDuplicidade: false }));
    let duplicateBlocks = [];
    let removed = new Set();

    if (els.removeDuplicatesToggle.checked) {
      const duplicates = findDuplicateBlocks(items, mapping.filterDefs);
      duplicateBlocks = duplicates.blocks;
      removed = duplicates.removed;
      items.forEach((item, idx) => { item.excluidaDuplicidade = removed.has(idx); });
    }

    const usedItems = items.filter((item) => !item.excluidaDuplicidade);
    const comparison = groupByName(usedItems, currentYear, previousYear, mapping.filterDefs);
    const periodSummary = groupByPeriod(usedItems, currentYear, previousYear);
    const categorySummary = groupByCategory(usedItems, currentYear, previousYear);
    const filterSummary = groupByFilters(usedItems, currentYear, previousYear, mapping.filterDefs);
    const totalCurrent = usedItems.filter((item) => item.ano === currentYear).reduce((acc, item) => acc + item.valorLiquido, 0);
    const totalPrevious = usedItems.filter((item) => item.ano === previousYear).reduce((acc, item) => acc + item.valorLiquido, 0);

    state.baseItems = baseItems;
    state.analysis = {
      currentYear,
      previousYear,
      detectedYears: years,
      periodMode,
      periodLabel: PERIOD_LABELS[periodMode] || periodMode,
      mapping,
      allItems: items,
      usedItems,
      comparison,
      periodSummary,
      categorySummary,
      filterSummary,
      duplicateBlocks,
      totals: {
        originalRows: baseItems.length,
        usedRows: usedItems.length,
        duplicatesRemoved: removed.size,
        names: new Set(comparison.map((row) => row.Nome_Extraido)).size,
        current: round2(totalCurrent),
        previous: round2(totalPrevious),
        variation: round2(totalCurrent - totalPrevious),
      },
    };

    renderAnalysis();
  }

  function renderAnalysis() {
    const analysis = state.analysis;
    if (!analysis) return;

    els.metricComparison.textContent = `${analysis.currentYear} x ${analysis.previousYear}`;
    els.metricPeriodMode.textContent = analysis.periodLabel;
    els.metricUsed.textContent = formatNumber(analysis.totals.usedRows);
    els.metricDuplicates.textContent = formatNumber(analysis.totals.duplicatesRemoved);
    els.metricNames.textContent = formatNumber(analysis.totals.names);
    els.metricVariation.textContent = formatMoney(analysis.totals.variation);

    els.summarySection.classList.remove('hidden');
    els.previewSection.classList.remove('hidden');
    els.exportBtn.disabled = false;
    els.resetBtn.disabled = false;

    renderFilterControls();
    renderPreviewTable();

    const duplicateText = analysis.totals.duplicatesRemoved
      ? `${formatNumber(analysis.totals.duplicatesRemoved)} linhas removidas por duplicidade de bloco.`
      : 'Nenhum bloco duplicado removido.';
    setStatus(
      `Análise concluída. Comparativo automático ${analysis.currentYear} x ${analysis.previousYear}, visão ${analysis.periodLabel}. ${duplicateText}`,
      'success',
    );
  }

  function renderFilterControls() {
    const analysis = state.analysis;
    const filterDefs = analysis.mapping.filterDefs;
    els.filterControls.innerHTML = '';
    if (!filterDefs.length) {
      els.filterControls.classList.add('hidden');
      return;
    }

    filterDefs.forEach((def) => {
      const label = document.createElement('label');
      label.textContent = def.label;
      const select = document.createElement('select');
      select.dataset.filterName = def.label;
      const allOpt = document.createElement('option');
      allOpt.value = '';
      allOpt.textContent = 'Todos';
      select.appendChild(allOpt);

      const values = [...new Set(analysis.usedItems.map((item) => item.filterValues[def.label]).filter(Boolean))].sort();
      values.forEach((value) => {
        const opt = document.createElement('option');
        opt.value = value;
        opt.textContent = value;
        select.appendChild(opt);
      });
      select.addEventListener('change', renderPreviewTable);
      label.appendChild(select);
      els.filterControls.appendChild(label);
    });
    els.filterControls.classList.remove('hidden');
  }

  function getActivePreviewFilters() {
    return Array.from(els.filterControls.querySelectorAll('select'))
      .map((select) => ({ name: select.dataset.filterName, value: select.value }))
      .filter((entry) => entry.value);
  }

  function renderPreviewTable() {
    const analysis = state.analysis;
    if (!analysis) return;
    const query = normalizeText(els.searchInput.value);
    const activeFilters = getActivePreviewFilters();
    const filterDefs = analysis.mapping.filterDefs;

    const rows = analysis.comparison
      .filter((row) => {
        if (activeFilters.some((filter) => !normalizeText(row[filter.name]).includes(normalizeText(filter.value)))) return false;
        if (!query) return true;
        const searchable = [row.Periodo, row.Nome_Extraido, row.Status, ...filterDefs.map((def) => row[def.label])].join(' | ');
        return normalizeText(searchable).includes(query);
      })
      .slice(0, 25);

    const headers = [
      'Periodo',
      'Nome_Extraido',
      ...filterDefs.map((def) => def.label),
      `Qtde_${analysis.previousYear}`,
      `Qtde_${analysis.currentYear}`,
      `Valor_${analysis.previousYear}`,
      `Valor_${analysis.currentYear}`,
      'Variacao_R$',
      'Variacao_%',
      'Status',
    ];

    els.previewTable.querySelector('thead').innerHTML = `<tr>${headers.map((h) => `<th>${escapeHTML(h)}</th>`).join('')}</tr>`;
    els.previewTable.querySelector('tbody').innerHTML = rows.map((row) => {
      return `<tr>${headers.map((header) => {
        const value = row[header];
        const isNumeric = /^Qtde_|^Valor_|Variacao/.test(header);
        let display = value;
        if (/^Valor_|Variacao_R\$/.test(header)) display = formatMoney(value);
        if (header === 'Variacao_%') display = formatPercent(value);
        return `<td class="${isNumeric ? 'numeric' : ''}">${escapeHTML(display ?? '-')}</td>`;
      }).join('')}</tr>`;
    }).join('');
  }

  function escapeHTML(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  async function handleFile(file) {
    if (!file) return;
    resetState(false);
    state.workbookName = file.name.replace(/\.[^.]+$/, '');
    setStatus('Lendo arquivo...', 'muted');

    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    state.workbook = workbook;
    state.rowsBySheet.clear();

    els.sheetSelect.innerHTML = '';
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
      state.rowsBySheet.set(sheetName, rows);
      const option = document.createElement('option');
      option.value = sheetName;
      option.textContent = sheetName;
      els.sheetSelect.appendChild(option);
    });

    els.sheetSelect.disabled = workbook.SheetNames.length < 2;
    els.resetBtn.disabled = false;
    if (!workbook.SheetNames.length) throw new Error('O arquivo não contém abas legíveis.');
    loadSheet(workbook.SheetNames[0]);
  }

  function toBaseExportRows(items, filterDefs) {
    return items.map((item) => {
      const row = {
        Linha_Original: item.linhaOriginal,
        Data: formatDateBR(item.data),
        Ano: item.ano,
        Mes: item.mes,
        Periodo: item.periodo,
        Categoria_Tratada: item.categoriaTratada,
        Nome_Extraido: item.nomeExtraido,
        NF: item.nf,
      };
      filterDefs.forEach((def) => { row[def.label] = item.filterValues[def.label] || ''; });
      return {
        ...row,
        Historico: item.historico,
        Debito: round2(item.debito),
        Credito: round2(item.credito),
        Valor_Liquido: round2(item.valorLiquido),
        Metodo_Extracao: item.metodoExtracao,
      };
    });
  }

  function selectedColumnName(index) {
    if (index === null || index === undefined || index === '') return 'Não usado';
    const def = state.columnDefs.find((col) => col.index === Number(index));
    return def ? `${def.letter} - ${def.label}` : String(index);
  }

  function buildSummaryRows(analysis) {
    const variationPct = analysis.totals.previous !== 0 ? analysis.totals.variation / analysis.totals.previous : null;
    const mapping = analysis.mapping;
    const filterNames = mapping.filterDefs.map((def) => `${def.letter} - ${def.label}`).join(', ') || 'Nenhuma';

    const rows = [
      [`Comparativo ${analysis.currentYear} x ${analysis.previousYear} por nome extraído do histórico`],
      [],
      ['Arquivo analisado', state.workbookName || 'Arquivo importado'],
      ['Aba analisada', els.sheetSelect.value],
      ['Visão de período', analysis.periodLabel],
      ['Anos detectados', analysis.detectedYears.join(', ')],
      ['Anos comparados automaticamente', `${analysis.currentYear} x ${analysis.previousYear}`],
      ['Coluna de data', selectedColumnName(mapping.data)],
      ['Coluna de histórico', selectedColumnName(mapping.historico)],
      ['Coluna de débito', selectedColumnName(mapping.debito)],
      ['Coluna de crédito', selectedColumnName(mapping.credito)],
      ['Coluna de valor pronto', selectedColumnName(mapping.valor)],
      ['Colunas de filtros adicionais', filterNames],
      ['Linhas originais lidas', analysis.totals.originalRows],
      ['Linhas usadas no comparativo', analysis.totals.usedRows],
      ['Linhas excluídas como duplicação de bloco', analysis.totals.duplicatesRemoved],
      ['Regra aplicada para autônomos', 'AUTONOMO + INSS S/ PF = AUTONOMO'],
      ['Critério do valor', mapping.valor === null ? 'Valor líquido = Débito - Crédito' : 'Valor líquido = coluna de valor pronto selecionada'],
      ['Nomes identificados', analysis.totals.names],
      [`Total ${analysis.previousYear}`, analysis.totals.previous],
      [`Total ${analysis.currentYear}`, analysis.totals.current],
      ['Variação R$', analysis.totals.variation],
      ['Variação %', variationPct],
      [],
      [`Top 15 por valor de ${analysis.currentYear}`],
      ['Período', 'Nome_Extraido', `Valor_${analysis.previousYear}`, `Valor_${analysis.currentYear}`, 'Variacao_R$', 'Status'],
    ];

    [...analysis.comparison]
      .sort((a, b) => Math.abs(b[`Valor_${analysis.currentYear}`] || 0) - Math.abs(a[`Valor_${analysis.currentYear}`] || 0))
      .slice(0, 15)
      .forEach((row) => {
        rows.push([
          row.Periodo,
          row.Nome_Extraido,
          row[`Valor_${analysis.previousYear}`],
          row[`Valor_${analysis.currentYear}`],
          row['Variacao_R$'],
          row.Status,
        ]);
      });

    return rows;
  }

  function hideInternalColumns(rows) {
    return rows.map((row) => {
      const copy = { ...row };
      delete copy.Periodo_Ordem;
      return copy;
    });
  }

  function jsonSheet(rows, columnsWidth = []) {
    const sheet = XLSX.utils.json_to_sheet(rows);
    if (columnsWidth.length) sheet['!cols'] = columnsWidth.map((wch) => ({ wch }));
    return sheet;
  }

  function aoaSheet(rows, columnsWidth = []) {
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    if (columnsWidth.length) sheet['!cols'] = columnsWidth.map((wch) => ({ wch }));
    return sheet;
  }

  function exportWorkbook() {
    const analysis = state.analysis;
    if (!analysis) return;

    const filterCount = analysis.mapping.filterDefs.length;
    const wb = XLSX.utils.book_new();
    const summarySheet = aoaSheet(buildSummaryRows(analysis), [54, 38, 18, 18, 18, 24]);
    const comparisonSheet = jsonSheet(hideInternalColumns(analysis.comparison), [18, 48, ...Array(filterCount).fill(28), 12, 12, 16, 16, 16, 14, 16, 24]);
    const periodSheet = jsonSheet(hideInternalColumns(analysis.periodSummary), [18, 12, 12, 16, 16, 16, 14]);
    const categorySheet = jsonSheet(hideInternalColumns(analysis.categorySummary), [18, 28, 12, 12, 16, 16, 16, 14]);
    const filterSheet = jsonSheet(hideInternalColumns(analysis.filterSummary), [28, 36, 18, 12, 12, 16, 16, 16, 14]);
    const baseSheet = jsonSheet(toBaseExportRows(analysis.usedItems, analysis.mapping.filterDefs), [14, 12, 10, 8, 18, 24, 48, 14, ...Array(filterCount).fill(30), 80, 14, 14, 16, 20]);
    const duplicatesSheet = jsonSheet(analysis.duplicateBlocks, [24, 24, 24, 24, 18]);

    XLSX.utils.book_append_sheet(wb, summarySheet, 'Resumo');
    XLSX.utils.book_append_sheet(wb, comparisonSheet, 'Comparativo_Nome');
    XLSX.utils.book_append_sheet(wb, periodSheet, 'Resumo_Periodo');
    XLSX.utils.book_append_sheet(wb, categorySheet, 'Resumo_Categoria');
    XLSX.utils.book_append_sheet(wb, filterSheet, 'Resumo_Filtros');
    XLSX.utils.book_append_sheet(wb, baseSheet, 'Base_Tratada');
    XLSX.utils.book_append_sheet(wb, duplicatesSheet, 'Duplicidades');

    const safeName = (state.workbookName || 'comparativo')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9_-]+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '') || 'comparativo';

    XLSX.writeFile(wb, `${safeName}_comparativo_${analysis.currentYear}x${analysis.previousYear}_${analysis.periodMode}.xlsx`);
  }

  function resetState(resetInput = true) {
    state.workbook = null;
    state.workbookName = '';
    state.rowsBySheet.clear();
    state.rawRows = [];
    state.headerRowIndex = -1;
    state.columnDefs = [];
    state.defaultColumns = null;
    state.baseItems = [];
    state.analysis = null;

    if (resetInput) els.fileInput.value = '';
    els.sheetSelect.innerHTML = '';
    els.dateColumnSelect.innerHTML = '';
    els.historyColumnSelect.innerHTML = '';
    els.debitColumnSelect.innerHTML = '';
    els.creditColumnSelect.innerHTML = '';
    els.valueColumnSelect.innerHTML = '';
    els.filterColumnsSelect.innerHTML = '';
    els.sheetSelect.disabled = true;
    enableMappingControls(false);
    els.analyzeBtn.disabled = true;
    els.exportBtn.disabled = true;
    els.resetBtn.disabled = true;
    els.summarySection.classList.add('hidden');
    els.previewSection.classList.add('hidden');
    els.filterControls.classList.add('hidden');
    els.filterControls.innerHTML = '';
    els.searchInput.value = '';
    setStatus('Nenhum arquivo carregado.', 'muted');
  }

  ['change'].forEach((eventName) => {
    els.dateColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.historyColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.debitColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.creditColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.valueColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.filterColumnsSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.periodModeSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
  });

  els.fileInput.addEventListener('change', (event) => {
    const [file] = event.target.files;
    handleFile(file).catch((error) => setStatus(`Erro ao ler arquivo: ${error.message}`, 'error'));
  });

  els.dropZone.addEventListener('dragover', (event) => {
    event.preventDefault();
    els.dropZone.classList.add('dragover');
  });

  els.dropZone.addEventListener('dragleave', () => {
    els.dropZone.classList.remove('dragover');
  });

  els.dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    els.dropZone.classList.remove('dragover');
    const [file] = event.dataTransfer.files;
    handleFile(file).catch((error) => setStatus(`Erro ao ler arquivo: ${error.message}`, 'error'));
  });

  els.sheetSelect.addEventListener('change', () => loadSheet(els.sheetSelect.value));

  els.analyzeBtn.addEventListener('click', () => {
    try {
      analyze();
    } catch (error) {
      setStatus(`Erro na análise: ${error.message}`, 'error');
    }
  });

  els.exportBtn.addEventListener('click', exportWorkbook);
  els.resetBtn.addEventListener('click', () => resetState(true));
  els.searchInput.addEventListener('input', renderPreviewTable);
  els.sampleRulesBtn.addEventListener('click', () => els.rulesDialog.showModal());
})();
