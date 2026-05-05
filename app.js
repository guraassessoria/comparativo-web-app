/* global XLSX */
(() => {
  const state = {
    workbook: null,
    workbooks: [],
    workbookName: '',
    rowsBySheet: new Map(),
    sourceDefs: [],
    activeSourceKey: '',
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
    sourceSheetsBlock: document.getElementById('sourceSheetsBlock'),
    sourceSheetsSelect: document.getElementById('sourceSheetsSelect'),
    columnPreviewBlock: document.getElementById('columnPreviewBlock'),
    columnPreviewTable: document.getElementById('columnPreviewTable'),
    selectAllSourcesBtn: document.getElementById('selectAllSourcesBtn'),
    clearSourcesBtn: document.getElementById('clearSourcesBtn'),
    analysisModeSelect: document.getElementById('analysisModeSelect'),
    periodModeSelect: document.getElementById('periodModeSelect'),
    dateColumnSelect: document.getElementById('dateColumnSelect'),
    historyColumnSelect: document.getElementById('historyColumnSelect'),
    debitColumnSelect: document.getElementById('debitColumnSelect'),
    creditColumnSelect: document.getElementById('creditColumnSelect'),
    valueColumnSelect: document.getElementById('valueColumnSelect'),
    filterColumnsSelect: document.getElementById('filterColumnsSelect'),
    clearFilterColumnsBtn: document.getElementById('clearFilterColumnsBtn'),
    extractionFieldsSelect: document.getElementById('extractionFieldsSelect'),
    customExtractionsInput: document.getElementById('customExtractionsInput'),
    selectDefaultExtractionsBtn: document.getElementById('selectDefaultExtractionsBtn'),
    clearExtractionsBtn: document.getElementById('clearExtractionsBtn'),
    removeDuplicatesToggle: document.getElementById('removeDuplicatesToggle'),
    reconciliationBlock: document.getElementById('reconciliationBlock'),
    reconciliationScopeSelect: document.getElementById('reconciliationScopeSelect'),
    reconciliationMatchModeSelect: document.getElementById('reconciliationMatchModeSelect'),
    reconciliationKeysSelect: document.getElementById('reconciliationKeysSelect'),
    selectDefaultReconciliationKeysBtn: document.getElementById('selectDefaultReconciliationKeysBtn'),
    clearReconciliationKeysBtn: document.getElementById('clearReconciliationKeysBtn'),
    sideColumnSelect: document.getElementById('sideColumnSelect'),
    toleranceInput: document.getElementById('toleranceInput'),
    requireOppositeSignsToggle: document.getElementById('requireOppositeSignsToggle'),
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
    metricInvoices: document.getElementById('metricInvoices'),
    metricExtractions: document.getElementById('metricExtractions'),
    metricVariation: document.getElementById('metricVariation'),
    metricResultLabel: document.getElementById('metricResultLabel'),
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

  const BUILTIN_EXTRACTION_LABELS = {
    Nota_Fiscal: 'Nota Fiscal',
    CNPJ_CPF: 'CNPJ/CPF',
    Competencia: 'Competência',
    Contrato: 'Contrato',
    Pedido_OC: 'Pedido/OC',
    Parcela: 'Parcela',
    Documento: 'Documento',
  };

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

  function formatSampleValue(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return formatDateBR(value);
    const text = String(value ?? '').replace(/\s+/g, ' ').trim();
    if (!text) return '';
    return text.length > 28 ? `${text.slice(0, 28)}...` : text;
  }

  function sampleColumnValues(rows, index, headerRowIndex) {
    const start = headerRowIndex >= 0 ? headerRowIndex + 1 : 0;
    const samples = [];
    for (let i = start; i < rows.length && samples.length < 3; i += 1) {
      const text = formatSampleValue(rows[i]?.[index]);
      if (text) samples.push(text);
    }
    return samples;
  }

  function compactColumnLabel(value, fallback) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return fallback;

    let text = String(value ?? '')
      .replace(/[\r\n\t]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    if (!text) return fallback;

    // Remove resíduos de versões anteriores, quando o rótulo trazia amostras concatenadas por "|".
    text = text.split('|')[0].trim();

    // Se a célula de cabeçalho parece um valor de linha, não a use como rótulo visual.
    if (/^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$/.test(text)) return fallback;
    if (/^[A-Z][a-z]{2} [A-Z][a-z]{2} \d{1,2} \d{4}/.test(text)) return fallback;
    if (/^\d{4}-\d{1,2}-\d{1,2}/.test(text)) return fallback;
    if (text.length > 34) text = `${text.slice(0, 31)}...`;

    return text || fallback;
  }

  function buildColumnDefs(rows, headerRowIndex) {
    const maxCols = getMaxColumnCount(rows);
    const headerRow = headerRowIndex >= 0 ? rows[headerRowIndex] || [] : [];
    const usedLabels = new Map();

    return Array.from({ length: maxCols }, (_, index) => {
      const letter = colLetter(index);
      const rawHeader = headerRow[index];
      const samples = sampleColumnValues(rows, index, headerRowIndex);
      let label = compactColumnLabel(rawHeader, `Coluna ${letter}`);
      const count = usedLabels.get(label) || 0;
      usedLabels.set(label, count + 1);
      const uniqueLabel = count ? `${label} (${letter})` : label;
      const sampleText = samples.join(' | ');
      const optionText = `${letter} — ${uniqueLabel}`;
      return { index, letter, label: uniqueLabel, optionText, sampleText };
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
      if (def.sampleText) option.title = `Amostras: ${def.sampleText}`;
      select.appendChild(option);
    });
  }

  function renderColumnPreview(columnDefs) {
    if (!els.columnPreviewBlock || !els.columnPreviewTable) return;
    const tbody = els.columnPreviewTable.querySelector('tbody');
    if (!columnDefs.length) {
      els.columnPreviewBlock.classList.add('hidden');
      if (tbody) tbody.innerHTML = '';
      return;
    }

    const rows = columnDefs.map((def) => {
      const samples = def.sampleText || 'Sem amostra';
      return `
        <tr>
          <td><strong>${escapeHTML(def.letter)}</strong></td>
          <td>${escapeHTML(def.optionText)}</td>
          <td>${escapeHTML(samples)}</td>
        </tr>`;
    }).join('');

    if (tbody) tbody.innerHTML = rows;
    els.columnPreviewBlock.classList.remove('hidden');
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
      els.analysisModeSelect,
      els.periodModeSelect,
      els.dateColumnSelect,
      els.historyColumnSelect,
      els.debitColumnSelect,
      els.creditColumnSelect,
      els.valueColumnSelect,
      els.filterColumnsSelect,
      els.clearFilterColumnsBtn,
      els.sourceSheetsSelect,
      els.selectAllSourcesBtn,
      els.clearSourcesBtn,
      els.extractionFieldsSelect,
      els.customExtractionsInput,
      els.selectDefaultExtractionsBtn,
      els.clearExtractionsBtn,
      els.reconciliationScopeSelect,
      els.reconciliationMatchModeSelect,
      els.reconciliationKeysSelect,
      els.selectDefaultReconciliationKeysBtn,
      els.clearReconciliationKeysBtn,
      els.sideColumnSelect,
      els.toleranceInput,
      els.requireOppositeSignsToggle,
    ].filter(Boolean).forEach((el) => { el.disabled = !enabled; });
    updateAnalysisModeUI();
  }

  function getSourceByKey(key) {
    return state.sourceDefs.find((source) => source.key === key) || null;
  }

  function getActiveSourceDef() {
    return getSourceByKey(els.sheetSelect.value) || state.sourceDefs[0] || null;
  }

  function selectedSourceDefsForAnalysis(mapping = null) {
    const analysisMode = mapping?.analysisMode || els.analysisModeSelect?.value || 'comparison';
    const reconciliationScope = mapping?.reconciliationScope || els.reconciliationScopeSelect?.value || 'single_file';

    if (analysisMode === 'reconciliation' && reconciliationScope === 'between_files') {
      const selectedKeys = Array.from(els.sourceSheetsSelect?.selectedOptions || []).map((option) => option.value);
      const selected = selectedKeys.map(getSourceByKey).filter(Boolean);
      return selected.length ? selected : state.sourceDefs.slice();
    }

    const active = getActiveSourceDef();
    return active ? [active] : [];
  }

  function refreshSourceSelectionDefaults() {
    if (!els.sourceSheetsSelect) return;
    const activeKey = els.sheetSelect.value || state.sourceDefs[0]?.key || '';
    Array.from(els.sourceSheetsSelect.options).forEach((option) => {
      option.selected = state.sourceDefs.length > 1 ? true : option.value === activeKey;
    });
  }

  function updateReconciliationScopeUI() {
    const isReconciliation = els.analysisModeSelect?.value === 'reconciliation';
    const isBetweenFiles = els.reconciliationScopeSelect?.value === 'between_files';
    const showSources = Boolean(isReconciliation && isBetweenFiles);
    if (els.sourceSheetsBlock) els.sourceSheetsBlock.classList.toggle('hidden', !showSources);
    if (els.sourceSheetsSelect) els.sourceSheetsSelect.disabled = !showSources || state.sourceDefs.length < 1;
    if (els.selectAllSourcesBtn) els.selectAllSourcesBtn.disabled = !showSources || state.sourceDefs.length < 1;
    if (els.clearSourcesBtn) els.clearSourcesBtn.disabled = !showSources || state.sourceDefs.length < 1;

    if (showSources && els.sourceSheetsSelect && !Array.from(els.sourceSheetsSelect.selectedOptions).length) {
      refreshSourceSelectionDefaults();
    }
  }

  function loadSheet(sourceKey) {
    const source = getSourceByKey(sourceKey) || state.sourceDefs[0];
    if (!source) return;

    state.activeSourceKey = source.key;
    state.rawRows = source.rows || [];
    state.headerRowIndex = source.headerRowIndex;
    state.columnDefs = source.columnDefs || [];
    state.baseItems = [];
    state.analysis = null;

    const defaults = source.defaults || detectDefaultColumns(state.rawRows, state.headerRowIndex, state.columnDefs);
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
    if (els.sideColumnSelect) addColumnOptions(els.sideColumnSelect, state.columnDefs, true, 'Não usar origem/lado');
    renderColumnPreview(state.columnDefs);

    setSelectValue(els.dateColumnSelect, defaults.data, fallback.data);
    setSelectValue(els.historyColumnSelect, defaults.historico, fallback.historico);
    setSelectValue(els.debitColumnSelect, defaults.debito, fallback.debito);
    setSelectValue(els.creditColumnSelect, defaults.credito, fallback.credito);
    setSelectValue(els.valueColumnSelect, defaults.valor, null);
    setMultiSelectValues(els.filterColumnsSelect, []);
    if (els.sideColumnSelect) els.sideColumnSelect.value = '';
    updateReconciliationKeyOptions(true);
    updateReconciliationScopeUI();

    enableMappingControls(state.columnDefs.length > 0);
    els.analyzeBtn.disabled = state.columnDefs.length === 0;
    els.exportBtn.disabled = true;
    els.summarySection.classList.add('hidden');
    els.previewSection.classList.add('hidden');
    els.filterControls.classList.add('hidden');
    els.filterControls.innerHTML = '';

    const headerInfo = state.headerRowIndex >= 0
      ? `Cabeçalho detectado na linha ${state.headerRowIndex + 1}. Confira as colunas antes de analisar.`
      : 'Sem cabeçalho detectado; as opções foram montadas apenas com letras de coluna para evitar rótulos confusos.';
    setStatus(`${formatNumber(state.rawRows.length)} linhas lidas da fonte selecionada (${source.label}). ${headerInfo}`, 'success');
  }

  function getSelectedFilterDefs() {
    const selected = Array.from(els.filterColumnsSelect.selectedOptions).map((option) => Number(option.value));
    return selected
      .filter((idx) => Number.isInteger(idx) && idx >= 0)
      .map((idx) => state.columnDefs.find((def) => def.index === idx))
      .filter(Boolean);
  }

  function sanitizeFieldName(name) {
    const cleaned = normalizeText(name)
      .replace(/[^A-Z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .slice(0, 40);
    return cleaned || 'CAMPO_EXTRAIDO';
  }

  function uniqueExtractionKey(base, existing) {
    let key = base;
    let count = 2;
    while (existing.has(key)) {
      key = `${base}_${count}`;
      count += 1;
    }
    existing.add(key);
    return key;
  }

  function parseCustomExtractionDefs() {
    const text = els.customExtractionsInput.value || '';
    const defs = [];
    const used = new Set(Object.keys(BUILTIN_EXTRACTION_LABELS));

    text.split(/\r?\n/).forEach((rawLine, index) => {
      const line = rawLine.trim();
      if (!line || line.startsWith('#')) return;
      const equalsIndex = line.indexOf('=');
      if (equalsIndex <= 0 || equalsIndex === line.length - 1) {
        throw new Error(`Regra personalizada inválida na linha ${index + 1}. Use NomeDoCampo=expressão regular.`);
      }
      const label = line.slice(0, equalsIndex).trim();
      const pattern = line.slice(equalsIndex + 1).trim();
      try {
        const regex = new RegExp(pattern, 'i');
        const key = uniqueExtractionKey(sanitizeFieldName(label), used);
        defs.push({ key, label: normalizeText(label), type: 'custom', regex });
      } catch (error) {
        throw new Error(`Expressão regular inválida na regra personalizada ${label}: ${error.message}`);
      }
    });

    return defs;
  }

  function getSelectedExtractionDefs() {
    const selectedBuiltins = Array.from(els.extractionFieldsSelect.selectedOptions)
      .map((option) => option.value)
      .filter((key) => BUILTIN_EXTRACTION_LABELS[key])
      .map((key) => ({ key, label: BUILTIN_EXTRACTION_LABELS[key], type: 'builtin' }));
    return [...selectedBuiltins, ...parseCustomExtractionDefs()];
  }

  function updateAnalysisModeUI() {
    if (!els.analysisModeSelect || !els.reconciliationBlock) return;
    const isReconciliation = els.analysisModeSelect.value === 'reconciliation';
    const isBetweenFiles = els.reconciliationScopeSelect?.value === 'between_files';
    let matchMode = els.reconciliationMatchModeSelect?.value || 'balance';
    const hasColumns = state.columnDefs.length > 0;

    if (isReconciliation && isBetweenFiles && els.reconciliationMatchModeSelect && matchMode === 'balance') {
      els.reconciliationMatchModeSelect.value = 'source_totals';
      matchMode = 'source_totals';
      if (els.requireOppositeSignsToggle) els.requireOppositeSignsToggle.checked = false;
    }

    els.reconciliationBlock.classList.toggle('hidden', !isReconciliation);
    if (els.reconciliationScopeSelect) els.reconciliationScopeSelect.disabled = !hasColumns || !isReconciliation;
    if (els.reconciliationMatchModeSelect) els.reconciliationMatchModeSelect.disabled = !hasColumns || !isReconciliation;
    if (els.requireOppositeSignsToggle) {
      els.requireOppositeSignsToggle.disabled = !hasColumns || !isReconciliation || matchMode === 'source_totals';
    }
    updateReconciliationScopeUI();
  }

  function reconciliationKeyOptions(filterDefs = null, extractionDefs = null) {
    const filters = filterDefs || getSelectedFilterDefs();
    let extractions = [];
    try {
      extractions = extractionDefs || getSelectedExtractionDefs();
    } catch (error) {
      extractions = [];
    }

    const options = [
      { value: 'FIELD::Nome_Extraido', label: 'Nome extraído' },
      { value: 'FIELD::Nota_Fiscal', label: 'Nota Fiscal' },
      { value: 'FIELD::Categoria_Tratada', label: 'Categoria tratada' },
      { value: 'FIELD::Origem_Lado', label: 'Origem/Lado' },
      { value: 'FIELD::Arquivo_Origem', label: 'Arquivo de origem' },
      { value: 'FIELD::Aba_Origem', label: 'Aba de origem' },
      { value: 'FIELD::Ano', label: 'Ano' },
      { value: 'FIELD::Periodo', label: 'Período' },
    ];

    filters.forEach((def) => {
      options.push({ value: `FILTER::${def.label}`, label: `Filtro: ${def.label}` });
    });

    extractions.forEach((def) => {
      if (def.key === 'Nota_Fiscal') return;
      options.push({ value: `EXT::${def.key}`, label: `Extração: ${def.label}` });
    });

    return options;
  }

  function setDefaultReconciliationKeys() {
    if (!els.reconciliationKeysSelect) return;
    const available = new Set(Array.from(els.reconciliationKeysSelect.options).map((option) => option.value));
    const defaults = ['FIELD::Nota_Fiscal', 'FIELD::Nome_Extraido'].filter((value) => available.has(value));
    Array.from(els.reconciliationKeysSelect.options).forEach((option) => {
      option.selected = defaults.includes(option.value);
    });
  }

  function updateReconciliationKeyOptions(forceDefaults = false) {
    if (!els.reconciliationKeysSelect) return;
    const selectedBefore = new Set(Array.from(els.reconciliationKeysSelect.selectedOptions).map((option) => option.value));
    const options = reconciliationKeyOptions();
    els.reconciliationKeysSelect.innerHTML = '';
    options.forEach((entry) => {
      const option = document.createElement('option');
      option.value = entry.value;
      option.textContent = entry.label;
      option.selected = selectedBefore.has(entry.value);
      els.reconciliationKeysSelect.appendChild(option);
    });
    if (forceDefaults || !Array.from(els.reconciliationKeysSelect.selectedOptions).length) setDefaultReconciliationKeys();
  }

  function getSelectedReconciliationKeys() {
    if (!els.reconciliationKeysSelect) return [];
    const optionMap = new Map(reconciliationKeyOptions().map((entry) => [entry.value, entry.label]));
    const selected = Array.from(els.reconciliationKeysSelect.selectedOptions)
      .map((option) => ({ value: option.value, label: optionMap.get(option.value) || option.textContent || option.value }));
    if (selected.length) return selected;
    return [
      { value: 'FIELD::Nota_Fiscal', label: 'Nota Fiscal' },
      { value: 'FIELD::Nome_Extraido', label: 'Nome extraído' },
    ];
  }

  function getSelectedMapping() {
    const data = Number(els.dateColumnSelect.value);
    const historico = Number(els.historyColumnSelect.value);
    const debito = els.debitColumnSelect.value === '' ? null : Number(els.debitColumnSelect.value);
    const credito = els.creditColumnSelect.value === '' ? null : Number(els.creditColumnSelect.value);
    const valor = els.valueColumnSelect.value === '' ? null : Number(els.valueColumnSelect.value);
    const filterDefs = getSelectedFilterDefs();
    const extractionDefs = getSelectedExtractionDefs();
    const categoryIndex = state.defaultColumns?.categoria ?? null;
    const analysisMode = els.analysisModeSelect?.value || 'comparison';
    const reconciliationScope = els.reconciliationScopeSelect?.value || 'single_file';
    const reconciliationMatchMode = els.reconciliationMatchModeSelect?.value || (reconciliationScope === 'between_files' ? 'source_totals' : 'balance');
    const sideColumn = els.sideColumnSelect?.value === '' ? null : Number(els.sideColumnSelect?.value);
    const tolerance = Math.max(0, parseNumber(els.toleranceInput?.value ?? 0.01));
    const requireOppositeSigns = reconciliationMatchMode === 'source_totals' ? false : Boolean(els.requireOppositeSignsToggle?.checked);
    const reconciliationKeys = getSelectedReconciliationKeys();

    if (!Number.isInteger(data) || data < 0) throw new Error('Selecione uma coluna de data válida.');
    if (!Number.isInteger(historico) || historico < 0) throw new Error('Selecione uma coluna de histórico válida.');
    if (valor === null && debito === null && credito === null) {
      throw new Error('Selecione uma coluna de valor pronto ou pelo menos uma coluna de débito/crédito.');
    }
    if (analysisMode === 'reconciliation' && !reconciliationKeys.length) {
      throw new Error('Selecione pelo menos uma chave de conciliação.');
    }
    if (analysisMode === 'reconciliation' && reconciliationScope === 'between_files') {
      const selectedSources = Array.from(els.sourceSheetsSelect?.selectedOptions || []);
      if (selectedSources.length < 2) throw new Error('Para conciliação entre arquivos, selecione pelo menos duas fontes.');
    }

    return {
      analysisMode,
      reconciliationScope,
      reconciliationMatchMode,
      data,
      historico,
      debito,
      credito,
      valor,
      filterDefs,
      extractionDefs,
      categoryIndex,
      sideColumn,
      tolerance,
      requireOppositeSigns,
      reconciliationKeys,
    };
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

  function normalizeNFNumber(rawNumber) {
    const digits = String(rawNumber ?? '').replace(/\D/g, '');
    if (!digits) return '';
    const stripped = digits.replace(/^0+/, '');
    return stripped || '0';
  }

  function nfPatterns() {
    return [
      /\b(?:REF\.?\s*)?(?:NF|N\.F\.|NFE|NF-E|NFSE|NFS-E|NFS)\.?\s*(?:N[ºO°]?\.?|NUM\.?|NUMERO\.?|NO\.?)?\s*[:\-\/]?\s*([0-9][0-9.\-/]{0,18})\b/i,
      /\bNOTA\s+FISCAL(?:\s+ELETRONICA)?\s*(?:N[ºO°]?\.?|NUM\.?|NUMERO\.?|NO\.?)?\s*[:\-\/]?\s*([0-9][0-9.\-/]{0,18})\b/i,
      /\bN[ºO°]\.?\s*(?:DA\s+)?(?:NF|N\.F\.|NOTA\s+FISCAL)\s*[:\-\/]?\s*([0-9][0-9.\-/]{0,18})\b/i,
    ];
  }

  function extractNF(history) {
    const text = normalizeText(history);
    for (const pattern of nfPatterns()) {
      const match = text.match(pattern);
      const nf = match ? normalizeNFNumber(match[1]) : '';
      if (nf) return nf;
    }
    return '';
  }

  function extractDocumentId(history) {
    const text = normalizeText(history);
    const cnpj = text.match(/\b\d{2}\.?\d{3}\.?\d{3}\/?\d{4}-?\d{2}\b/);
    if (cnpj) return cnpj[0].replace(/\D/g, '').replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
    const cpf = text.match(/\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b/);
    if (cpf) return cpf[0].replace(/\D/g, '').replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
    return '';
  }

  function extractCompetencia(history) {
    const text = normalizeText(history);
    const numeric = text.match(/\b(?:COMP(?:ETENCIA)?|REF(?:ERENCIA)?|MES)\.?:?\s*(0?[1-9]|1[0-2])\s*[\/-]\s*(20\d{2}|\d{2})\b/i)
      || text.match(/\b(0?[1-9]|1[0-2])\s*[\/-]\s*(20\d{2})\b/i);
    if (numeric) {
      const month = String(Number(numeric[1])).padStart(2, '0');
      let year = String(numeric[2]);
      if (year.length === 2) year = `20${year}`;
      return `${month}/${year}`;
    }

    const monthMap = {
      JANEIRO: '01', FEVEREIRO: '02', MARCO: '03', MARÇO: '03', ABRIL: '04', MAIO: '05', JUNHO: '06',
      JULHO: '07', AGOSTO: '08', SETEMBRO: '09', OUTUBRO: '10', NOVEMBRO: '11', DEZEMBRO: '12',
    };
    const words = Object.keys(monthMap).join('|');
    const wordMatch = text.match(new RegExp(`\b(${words})\s*(?:DE\s*)?(20\d{2})\b`, 'i'));
    if (wordMatch) return `${monthMap[normalizeText(wordMatch[1])]}/${wordMatch[2]}`;
    return '';
  }

  function extractGenericCode(history, labels, maxLen = 30) {
    const text = normalizeText(history);
    const labelPattern = labels.join('|');
    const pattern = new RegExp(`\b(?:${labelPattern})\.?\s*(?:N[ºO°]?\.?|NUM(?:ERO)?\.?|NO\.?|:|-)?\s*([A-Z0-9][A-Z0-9._\/-]{0,${maxLen - 1}})`, 'i');
    const match = text.match(pattern);
    return match ? match[1].replace(/[.,;:]+$/g, '') : '';
  }

  function runExtractionDef(history, def) {
    const text = normalizeText(history);
    if (def.type === 'custom') {
      const match = text.match(def.regex);
      return match ? normalizeText(match[1] || match[0]) : '';
    }
    if (def.key === 'Nota_Fiscal') return extractNF(text);
    if (def.key === 'CNPJ_CPF') return extractDocumentId(text);
    if (def.key === 'Competencia') return extractCompetencia(text);
    if (def.key === 'Contrato') return extractGenericCode(text, ['CONTRATO', 'CONTR', 'CTR', 'CT']);
    if (def.key === 'Pedido_OC') return extractGenericCode(text, ['PEDIDO', 'PED', 'OC', 'ORDEM\s+DE\s+COMPRA']);
    if (def.key === 'Parcela') return extractGenericCode(text, ['PARCELA', 'PARC', 'PCL']);
    if (def.key === 'Documento') return extractGenericCode(text, ['DOCUMENTO', 'DOC']);
    return '';
  }

  function runExtractions(history, extractionDefs) {
    const out = {};
    (extractionDefs || []).forEach((def) => {
      out[def.key] = runExtractionDef(history, def);
    });
    return out;
  }

  function countExtractedValues(items, extractionDefs) {
    const unique = new Set();
    (items || []).forEach((item) => {
      (extractionDefs || []).forEach((def) => {
        const value = item.extracoes?.[def.key];
        if (value) unique.add(`${def.key}|${value}`);
      });
    });
    return unique.size;
  }

  function extractNameAfterNF(history) {
    const original = normalizeText(history);
    const patterns = [
      /\b(?:REF\.?\s*)?(?:NF|N\.F\.|NFE|NF-E|NFSE|NFS-E|NFS)\.?\s*(?:N[ºO°]?\.?|NUM\.?|NUMERO\.?|NO\.?)?\s*[:\-\/]?\s*[0-9][0-9.\-/]{0,18}\s*[-–—:]?\s*(.+)$/i,
      /\bNOTA\s+FISCAL(?:\s+ELETRONICA)?\s*(?:N[ºO°]?\.?|NUM\.?|NUMERO\.?|NO\.?)?\s*[:\-\/]?\s*[0-9][0-9.\-/]{0,18}\s*[-–—:]?\s*(.+)$/i,
      /\bN[ºO°]\.?\s*(?:DA\s+)?(?:NF|N\.F\.|NOTA\s+FISCAL)\s*[:\-\/]?\s*[0-9][0-9.\-/]{0,18}\s*[-–—:]?\s*(.+)$/i,
    ];
    for (const pattern of patterns) {
      const match = original.match(pattern);
      if (match) return cleanName(match[1]);
    }
    return '';
  }

  function extractNameFromRefNF(history) {
    const text = normalizeText(history);
    const isRefNF = /\bREF\.?\s+(?:NF|N\.F\.|NFE|NF-E|NFSE|NFS-E|NFS|NOTA\s+FISCAL)\b/i.test(text)
      || /^(?:NF|N\.F\.|NFE|NF-E|NFSE|NFS-E|NFS)\.?\s*(?:N[ºO°]?\.?\s*)?[0-9]/i.test(text)
      || /^NOTA\s+FISCAL/i.test(text);
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

  function rowToItem(row, sourceIndex, mapping, periodMode, source) {
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
    const sourceLabel = source?.label || '';
    const origemLado = Number.isInteger(mapping.sideColumn) && mapping.sideColumn >= 0
      ? normalizeText(row[mapping.sideColumn])
      : (mapping.analysisMode === 'reconciliation' && mapping.reconciliationScope === 'between_files' ? normalizeText(sourceLabel) : '');
    const period = periodMeta(date, periodMode);

    return {
      sourceKey: source?.key || '',
      sourceFile: source?.fileName || '',
      sourceSheet: source?.sheetName || '',
      sourceLabel,
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
      origemLado,
      historico,
      debito,
      credito,
      valorLiquido,
      nf: '',
      extracoes: {},
      nomeExtraido: '',
      metodoExtracao: '',
      excluidaDuplicidade: false,
    };
  }

  function buildBaseItems(sources, mapping, periodMode) {
    const sourceList = Array.isArray(sources) ? sources : [];
    const preliminary = [];

    sourceList.forEach((source) => {
      const rows = source.rows || [];
      const start = source.headerRowIndex >= 0 ? source.headerRowIndex + 1 : 0;
      const usableRows = rows.slice(start);

      usableRows.forEach((row, offset) => {
        const sourceIndex = start + offset;
        if (!row || row.every((cell) => cell === null || cell === undefined || String(cell).trim() === '')) return;
        const item = rowToItem(row, sourceIndex, mapping, periodMode, source);
        if (!item.historico && !item.data) return;
        preliminary.push(item);
      });
    });

    const nfMaps = buildNFMaps(preliminary);
    preliminary.forEach((item) => {
      const extracted = inferNameAndMethod(item, nfMaps);
      item.nf = extracted.nf;
      item.extracoes = runExtractions(item.historico, mapping.extractionDefs);
      if (Object.prototype.hasOwnProperty.call(item.extracoes, 'Nota_Fiscal') && !item.extracoes.Nota_Fiscal) {
        item.extracoes.Nota_Fiscal = item.nf;
      }
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

  function findDuplicateBlocksBySource(items, filterDefs, minRun = 20) {
    const removed = new Set();
    const blocks = [];
    const bySource = new Map();

    items.forEach((item, globalIndex) => {
      const key = item.sourceKey || '__single_source__';
      if (!bySource.has(key)) bySource.set(key, []);
      bySource.get(key).push({ item, globalIndex });
    });

    bySource.forEach((entries) => {
      const localItems = entries.map((entry) => entry.item);
      const result = findDuplicateBlocks(localItems, filterDefs, minRun);
      result.removed.forEach((localIndex) => removed.add(entries[localIndex].globalIndex));
      result.blocks.forEach((block) => {
        blocks.push({
          Fonte: entries[0]?.item?.sourceLabel || '',
          ...block,
        });
      });
    });

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

  function groupByNF(items, currentYear, previousYear, filterDefs) {
    const map = new Map();
    items.forEach((item) => {
      if (item.ano !== currentYear && item.ano !== previousYear) return;
      if (!item.nf) return;
      const key = `${item.periodoChave}|${item.nf}|${item.nomeExtraido || 'NAO IDENTIFICADO'}`;
      if (!map.has(key)) {
        map.set(key, {
          Periodo_Ordem: item.periodoOrdem,
          Periodo: item.periodo,
          Nota_Fiscal: item.nf,
          Nome_Extraido: item.nomeExtraido || 'NAO IDENTIFICADO',
          Categorias: new Set(),
          Filtros: {},
          [`Qtde_${previousYear}`]: 0,
          [`Qtde_${currentYear}`]: 0,
          [`Valor_${previousYear}`]: 0,
          [`Valor_${currentYear}`]: 0,
        });
      }
      const row = map.get(key);
      if (item.categoriaTratada) row.Categorias.add(item.categoriaTratada);
      addDistinctValues(row.Filtros, item.filterValues);
      row[`Qtde_${item.ano}`] += 1;
      row[`Valor_${item.ano}`] += item.valorLiquido;
    });

    return [...map.values()].map((row) => {
      const variacao = row[`Valor_${currentYear}`] - row[`Valor_${previousYear}`];
      const filterColumns = serializeFilterSets(row.Filtros);
      const output = {
        Periodo_Ordem: row.Periodo_Ordem,
        Periodo: row.Periodo,
        Nota_Fiscal: row.Nota_Fiscal,
        Nome_Extraido: row.Nome_Extraido,
        Categoria_Tratada: [...row.Categorias].sort().join(', '),
      };
      filterDefs.forEach((def) => { output[def.label] = filterColumns[def.label] || ''; });
      return {
        ...output,
        [`Qtde_${previousYear}`]: row[`Qtde_${previousYear}`],
        [`Qtde_${currentYear}`]: row[`Qtde_${currentYear}`],
        [`Valor_${previousYear}`]: round2(row[`Valor_${previousYear}`]),
        [`Valor_${currentYear}`]: round2(row[`Valor_${currentYear}`]),
        'Variacao_R$': round2(variacao),
        'Variacao_%': row[`Valor_${previousYear}`] !== 0 ? variacao / row[`Valor_${previousYear}`] : null,
      };
    }).sort((a, b) => {
      if ((a.Periodo_Ordem || 0) !== (b.Periodo_Ordem || 0)) return (a.Periodo_Ordem || 0) - (b.Periodo_Ordem || 0);
      const nfCompare = String(a.Nota_Fiscal || '').localeCompare(String(b.Nota_Fiscal || ''), 'pt-BR', { numeric: true });
      if (nfCompare !== 0) return nfCompare;
      return String(a.Nome_Extraido || '').localeCompare(String(b.Nome_Extraido || ''), 'pt-BR');
    });
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

  function groupByExtractions(items, currentYear, previousYear, extractionDefs) {
    const map = new Map();
    (extractionDefs || []).forEach((def) => {
      items.forEach((item) => {
        if (item.ano !== currentYear && item.ano !== previousYear) return;
        const value = item.extracoes?.[def.key];
        if (!value) return;
        const key = `${def.key}|${value}|${item.periodoChave}`;
        if (!map.has(key)) {
          map.set(key, {
            Campo_Extraido: def.label,
            Valor_Extraido: value,
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
      if (a.Campo_Extraido !== b.Campo_Extraido) return a.Campo_Extraido.localeCompare(b.Campo_Extraido);
      if ((a.Periodo_Ordem || 0) !== (b.Periodo_Ordem || 0)) return (a.Periodo_Ordem || 0) - (b.Periodo_Ordem || 0);
      return Math.abs(b[`Valor_${currentYear}`] || 0) - Math.abs(a[`Valor_${currentYear}`] || 0);
    });
  }

  function groupByExtractionAndName(items, currentYear, previousYear, extractionDefs) {
    const map = new Map();
    (extractionDefs || []).forEach((def) => {
      items.forEach((item) => {
        if (item.ano !== currentYear && item.ano !== previousYear) return;
        const value = item.extracoes?.[def.key];
        if (!value) return;
        const name = item.nomeExtraido || 'NAO IDENTIFICADO';
        const key = `${def.key}|${value}|${item.periodoChave}|${name}`;
        if (!map.has(key)) {
          map.set(key, {
            Campo_Extraido: def.label,
            Valor_Extraido: value,
            Periodo_Ordem: item.periodoOrdem,
            Periodo: item.periodo,
            Nome_Extraido: name,
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
      if (a.Campo_Extraido !== b.Campo_Extraido) return a.Campo_Extraido.localeCompare(b.Campo_Extraido);
      if ((a.Periodo_Ordem || 0) !== (b.Periodo_Ordem || 0)) return (a.Periodo_Ordem || 0) - (b.Periodo_Ordem || 0);
      const byValue = String(a.Valor_Extraido || '').localeCompare(String(b.Valor_Extraido || ''), 'pt-BR', { numeric: true });
      if (byValue !== 0) return byValue;
      return Math.abs(b[`Valor_${currentYear}`] || 0) - Math.abs(a[`Valor_${currentYear}`] || 0);
    });
  }

  function reconciliationKeyCell(item, keyDef) {
    const [kind, rawKey] = String(keyDef.value || '').split('::');
    if (kind === 'FIELD') {
      if (rawKey === 'Nome_Extraido') return item.nomeExtraido || 'NAO IDENTIFICADO';
      if (rawKey === 'Nota_Fiscal') return item.nf || item.extracoes?.Nota_Fiscal || 'SEM NF';
      if (rawKey === 'Categoria_Tratada') return item.categoriaTratada || 'SEM CATEGORIA';
      if (rawKey === 'Origem_Lado') return item.origemLado || 'SEM ORIGEM';
      if (rawKey === 'Arquivo_Origem') return item.sourceFile || 'SEM ARQUIVO';
      if (rawKey === 'Aba_Origem') return item.sourceSheet || 'SEM ABA';
      if (rawKey === 'Ano') return item.ano || 'SEM ANO';
      if (rawKey === 'Periodo') return item.periodo || 'SEM PERIODO';
    }
    if (kind === 'FILTER') return item.filterValues?.[rawKey] || 'VAZIO';
    if (kind === 'EXT') return item.extracoes?.[rawKey] || 'VAZIO';
    return 'VAZIO';
  }

  function reconciliationKeyObject(item, keyDefs) {
    const out = {};
    (keyDefs || []).forEach((keyDef) => {
      out[keyDef.label] = reconciliationKeyCell(item, keyDef);
    });
    return out;
  }

  function reconciliationKeySignature(item, keyDefs) {
    return (keyDefs || [])
      .map((keyDef) => `${keyDef.label}=${normalizeText(reconciliationKeyCell(item, keyDef))}`)
      .join('¦') || `Linha=${item.linhaOriginal}`;
  }

  function usesInvoiceAndNameKeys(keyDefs) {
    const selected = new Set((keyDefs || []).map((keyDef) => keyDef.value));
    return selected.has('FIELD::Nota_Fiscal') && selected.has('FIELD::Nome_Extraido');
  }

  function isConcreteInvoiceNameMatch(group, keyDefs) {
    if (!usesInvoiceAndNameKeys(keyDefs) || !group || group.Qtde_Linhas <= 1) return false;
    const nf = group._keys?.['Nota Fiscal'];
    const name = group._keys?.['Nome extraído'];
    if (!nf || !name) return false;
    return normalizeText(nf) !== 'SEM NF' && normalizeText(name) !== 'NAO IDENTIFICADO';
  }

  function observationForReconciliationStatus(status) {
    if (status === 'Divergência de valor') {
      return 'Nota fiscal e nome conferem, mas os valores divergem acima da tolerância.';
    }
    if (status === 'Conciliado') return 'Valores conciliados dentro da tolerância.';
    if (status === 'Conciliado - um lado') return 'Saldo zerado dentro de uma única origem/lado.';
    if (status === 'Sem contraparte') return 'Não há contraparte para a chave selecionada.';
    if (status === 'Sem sinais opostos') return 'Há match pela chave, mas os valores não têm sinais opostos.';
    if (status === 'Diferença') return 'Há diferença acima da tolerância para a chave selecionada.';
    return '';
  }

  function classifyReconciliationGroup(count, saldo, positiveTotal, negativeTotal, tolerance, requireOppositeSigns, sideCount) {
    if (count <= 1) return 'Sem contraparte';
    if (requireOppositeSigns && (Math.abs(positiveTotal) <= tolerance || Math.abs(negativeTotal) <= tolerance)) return 'Sem sinais opostos';
    if (Math.abs(saldo) <= tolerance) return sideCount === 1 ? 'Conciliado - um lado' : 'Conciliado';
    return 'Diferença';
  }

  function classifyBetweenFilesGroup(originTotals, tolerance) {
    const active = [...originTotals.entries()].filter(([, total]) => Math.abs(total || 0) > tolerance);
    if (active.length <= 1) return { status: 'Sem contraparte', difference: active.length ? Math.abs(active[0][1] || 0) : 0 };
    const absTotals = active.map(([, total]) => Math.abs(total || 0));
    const max = Math.max(...absTotals);
    const min = Math.min(...absTotals);
    const difference = max - min;
    return { status: difference <= tolerance ? 'Conciliado' : 'Diferença', difference };
  }

  function buildReconciliationGroups(items, mapping) {
    const keyDefs = mapping.reconciliationKeys || [];
    const groups = new Map();
    const betweenFiles = mapping.reconciliationScope === 'between_files' || mapping.reconciliationMatchMode === 'source_totals';

    items.forEach((item) => {
      const signature = reconciliationKeySignature(item, keyDefs);
      if (!groups.has(signature)) {
        groups.set(signature, {
          _signature: signature,
          _keys: reconciliationKeyObject(item, keyDefs),
          _linhas: [],
          _origens: new Set(),
          _originTotals: new Map(),
          _originCounts: new Map(),
          Qtde_Linhas: 0,
          Qtde_Debitos: 0,
          Qtde_Creditos: 0,
          Total_Debitos: 0,
          Total_Creditos: 0,
          Valor_Positivo: 0,
          Valor_Negativo: 0,
          Saldo: 0,
          Primeiro_Lancamento: '',
          Ultimo_Lancamento: '',
        });
      }
      const group = groups.get(signature);
      const origin = betweenFiles
        ? (item.sourceLabel || item.origemLado || 'SEM ORIGEM')
        : (item.origemLado || item.sourceLabel || 'SEM ORIGEM');
      group._linhas.push(item);
      if (origin) group._origens.add(origin);
      group._originTotals.set(origin, (group._originTotals.get(origin) || 0) + (item.valorLiquido || 0));
      group._originCounts.set(origin, (group._originCounts.get(origin) || 0) + 1);
      group.Qtde_Linhas += 1;
      group.Total_Debitos += item.debito || 0;
      group.Total_Creditos += item.credito || 0;
      if ((item.valorLiquido || 0) > 0) {
        group.Qtde_Debitos += 1;
        group.Valor_Positivo += item.valorLiquido || 0;
      } else if ((item.valorLiquido || 0) < 0) {
        group.Qtde_Creditos += 1;
        group.Valor_Negativo += item.valorLiquido || 0;
      }
      group.Saldo += item.valorLiquido || 0;
      const iso = item.dataISO || '';
      if (iso) {
        if (!group.Primeiro_Lancamento || iso < group.Primeiro_Lancamento) group.Primeiro_Lancamento = iso;
        if (!group.Ultimo_Lancamento || iso > group.Ultimo_Lancamento) group.Ultimo_Lancamento = iso;
      }
    });

    return [...groups.values()].map((group) => {
      const originTotalsText = [...group._originTotals.entries()]
        .sort((a, b) => a[0].localeCompare(b[0], 'pt-BR', { numeric: true }))
        .map(([origin, total]) => `${origin}: ${round2(total)}`)
        .join(' | ');
      const originCountText = [...group._originCounts.entries()]
        .sort((a, b) => a[0].localeCompare(b[0], 'pt-BR', { numeric: true }))
        .map(([origin, count]) => `${origin}: ${count}`)
        .join(' | ');

      let status;
      let differenceAbs;
      if (betweenFiles) {
        const classification = classifyBetweenFilesGroup(group._originTotals, mapping.tolerance);
        status = classification.status;
        differenceAbs = classification.difference;
      } else {
        status = classifyReconciliationGroup(
          group.Qtde_Linhas,
          group.Saldo,
          group.Valor_Positivo,
          group.Valor_Negativo,
          mapping.tolerance,
          mapping.requireOppositeSigns,
          group._origens.size,
        );
        differenceAbs = Math.abs(group.Saldo);
      }

      const matchedByInvoiceAndName = isConcreteInvoiceNameMatch(group, keyDefs);
      if (matchedByInvoiceAndName && Math.abs(differenceAbs || 0) > mapping.tolerance && status !== 'Sem contraparte' && status !== 'Sem sinais opostos') {
        status = 'Divergência de valor';
      }

      return {
        Chave_Conciliacao: group._signature,
        ...group._keys,
        Status_Conciliacao: status,
        Observacao_Conciliacao: observationForReconciliationStatus(status),
        Match_NF_Nome: matchedByInvoiceAndName ? 'Sim' : 'Não',
        Metodo_Conciliacao: betweenFiles ? 'Entre arquivos/abas' : 'Dentro de um arquivo/aba',
        Qtde_Linhas: group.Qtde_Linhas,
        Qtde_Debitos: group.Qtde_Debitos,
        Qtde_Creditos: group.Qtde_Creditos,
        Qtde_Origens: group._origens.size || '',
        Origens_Lados: [...group._origens].sort().join(', '),
        Total_Por_Origem: originTotalsText,
        Linhas_Por_Origem: originCountText,
        Total_Debitos: round2(group.Total_Debitos),
        Total_Creditos: round2(group.Total_Creditos),
        Valor_Positivo: round2(group.Valor_Positivo),
        Valor_Negativo: round2(group.Valor_Negativo),
        Saldo: round2(group.Saldo),
        Diferenca_Absoluta: round2(differenceAbs),
        Primeiro_Lancamento: group.Primeiro_Lancamento,
        Ultimo_Lancamento: group.Ultimo_Lancamento,
        _linhas: group._linhas,
        _originTotals: group._originTotals,
      };
    }).sort((a, b) => {
      const statusOrder = { 'Divergência de valor': 1, 'Diferença': 2, 'Sem contraparte': 3, 'Sem sinais opostos': 4, 'Conciliado - um lado': 5, Conciliado: 6 };
      const byStatus = (statusOrder[a.Status_Conciliacao] || 99) - (statusOrder[b.Status_Conciliacao] || 99);
      if (byStatus !== 0) return byStatus;
      return Math.abs(b.Diferenca_Absoluta || b.Saldo || 0) - Math.abs(a.Diferenca_Absoluta || a.Saldo || 0);
    });
  }

  function buildReconciliationLines(groups) {
    const rows = [];
    groups.forEach((group) => {
      group._linhas.forEach((item) => {
        rows.push({
          Chave_Conciliacao: group.Chave_Conciliacao,
          Status_Conciliacao: group.Status_Conciliacao,
          Observacao_Conciliacao: group.Observacao_Conciliacao || '',
          Match_NF_Nome: group.Match_NF_Nome || '',
          Saldo_Grupo: group.Saldo,
          Diferenca_Grupo: group.Diferenca_Absoluta,
          Arquivo_Origem: item.sourceFile || '',
          Aba_Origem: item.sourceSheet || '',
          Linha_Original: item.linhaOriginal,
          Data: formatDateBR(item.data),
          Ano: item.ano,
          Periodo: item.periodo,
          Origem_Lado: item.origemLado || '',
          Categoria_Tratada: item.categoriaTratada,
          Nome_Extraido: item.nomeExtraido,
          Nota_Fiscal: item.nf,
          Historico: item.historico,
          Debito: round2(item.debito),
          Credito: round2(item.credito),
          Valor_Liquido: round2(item.valorLiquido),
          Metodo_Extracao: item.metodoExtracao,
        });
      });
    });
    return rows;
  }

  function buildReconciliationPairsWithinFile(groups, tolerance) {
    const rows = [];
    groups.forEach((group) => {
      const positives = group._linhas
        .filter((item) => (item.valorLiquido || 0) > tolerance)
        .sort((a, b) => Math.abs(b.valorLiquido) - Math.abs(a.valorLiquido));
      const negatives = group._linhas
        .filter((item) => (item.valorLiquido || 0) < -tolerance)
        .sort((a, b) => Math.abs(b.valorLiquido) - Math.abs(a.valorLiquido));
      const usedNegatives = new Set();
      const usedPositives = new Set();

      positives.forEach((pos, posIndex) => {
        let bestIndex = -1;
        let bestDiff = Infinity;
        negatives.forEach((neg, negIndex) => {
          if (usedNegatives.has(negIndex)) return;
          const diff = Math.abs((pos.valorLiquido || 0) + (neg.valorLiquido || 0));
          if (diff < bestDiff) {
            bestIndex = negIndex;
            bestDiff = diff;
          }
        });
        if (bestIndex >= 0) {
          const neg = negatives[bestIndex];
          usedNegatives.add(bestIndex);
          usedPositives.add(posIndex);
          rows.push({
            Chave_Conciliacao: group.Chave_Conciliacao,
            Status_Par: bestDiff <= tolerance ? 'Par conciliado' : 'Verificar - valor divergente',
            Linha_Positiva: pos.linhaOriginal,
            Data_Positiva: formatDateBR(pos.data),
            Valor_Positivo: round2(pos.valorLiquido),
            Historico_Positivo: pos.historico,
            Linha_Negativa: neg.linhaOriginal,
            Data_Negativa: formatDateBR(neg.data),
            Valor_Negativo: round2(neg.valorLiquido),
            Historico_Negativo: neg.historico,
            Diferenca: round2((pos.valorLiquido || 0) + (neg.valorLiquido || 0)),
          });
        }
      });

      positives.forEach((pos, posIndex) => {
        if (usedPositives.has(posIndex)) return;
        rows.push({
          Chave_Conciliacao: group.Chave_Conciliacao,
          Status_Par: 'Positivo sem par',
          Linha_Positiva: pos.linhaOriginal,
          Data_Positiva: formatDateBR(pos.data),
          Valor_Positivo: round2(pos.valorLiquido),
          Historico_Positivo: pos.historico,
          Linha_Negativa: '',
          Data_Negativa: '',
          Valor_Negativo: '',
          Historico_Negativo: '',
          Diferenca: round2(pos.valorLiquido),
        });
      });

      negatives.forEach((neg, negIndex) => {
        if (usedNegatives.has(negIndex)) return;
        rows.push({
          Chave_Conciliacao: group.Chave_Conciliacao,
          Status_Par: 'Negativo sem par',
          Linha_Positiva: '',
          Data_Positiva: '',
          Valor_Positivo: '',
          Historico_Positivo: '',
          Linha_Negativa: neg.linhaOriginal,
          Data_Negativa: formatDateBR(neg.data),
          Valor_Negativo: round2(neg.valorLiquido),
          Historico_Negativo: neg.historico,
          Diferenca: round2(neg.valorLiquido),
        });
      });
    });
    return rows;
  }

  function buildReconciliationPairsBetweenFiles(groups, tolerance) {
    const rows = [];
    groups.forEach((group) => {
      const byOrigin = new Map();
      group._linhas.forEach((item) => {
        const origin = item.sourceLabel || item.origemLado || 'SEM ORIGEM';
        if (!byOrigin.has(origin)) byOrigin.set(origin, []);
        byOrigin.get(origin).push(item);
      });

      const origins = [...byOrigin.keys()].sort();
      if (origins.length < 2) {
        group._linhas.forEach((item) => {
          rows.push({
            Chave_Conciliacao: group.Chave_Conciliacao,
            Status_Par: 'Linha sem contraparte em outra origem',
            Origem_Referencia: item.sourceLabel || item.origemLado || '',
            Arquivo_Referencia: item.sourceFile || '',
            Linha_Referencia: item.linhaOriginal,
            Data_Referencia: formatDateBR(item.data),
            Valor_Referencia: round2(item.valorLiquido),
            Historico_Referencia: item.historico,
            Origem_Comparada: '',
            Arquivo_Comparado: '',
            Linha_Comparada: '',
            Data_Comparada: '',
            Valor_Comparado: '',
            Historico_Comparado: '',
            Diferenca: round2(Math.abs(item.valorLiquido || 0)),
          });
        });
        return;
      }

      const referenceOrigin = origins[0];
      const referenceRows = [...byOrigin.get(referenceOrigin)].sort((a, b) => Math.abs(b.valorLiquido) - Math.abs(a.valorLiquido));
      const otherOrigins = origins.slice(1);

      otherOrigins.forEach((origin) => {
        const candidates = [...byOrigin.get(origin)].sort((a, b) => Math.abs(b.valorLiquido) - Math.abs(a.valorLiquido));
        const usedCandidates = new Set();
        const usedReferences = new Set();

        referenceRows.forEach((ref, refIndex) => {
          let bestIndex = -1;
          let bestDiff = Infinity;
          candidates.forEach((candidate, candidateIndex) => {
            if (usedCandidates.has(candidateIndex)) return;
            const diff = Math.abs(Math.abs(ref.valorLiquido || 0) - Math.abs(candidate.valorLiquido || 0));
            if (diff < bestDiff) {
              bestIndex = candidateIndex;
              bestDiff = diff;
            }
          });
          if (bestIndex >= 0) {
            const cmp = candidates[bestIndex];
            usedCandidates.add(bestIndex);
            usedReferences.add(refIndex);
            rows.push({
              Chave_Conciliacao: group.Chave_Conciliacao,
              Status_Par: bestDiff <= tolerance ? 'Par conciliado entre origens' : 'Verificar - valor divergente',
              Origem_Referencia: referenceOrigin,
              Arquivo_Referencia: ref.sourceFile || '',
              Linha_Referencia: ref.linhaOriginal,
              Data_Referencia: formatDateBR(ref.data),
              Valor_Referencia: round2(ref.valorLiquido),
              Historico_Referencia: ref.historico,
              Origem_Comparada: origin,
              Arquivo_Comparado: cmp.sourceFile || '',
              Linha_Comparada: cmp.linhaOriginal,
              Data_Comparada: formatDateBR(cmp.data),
              Valor_Comparado: round2(cmp.valorLiquido),
              Historico_Comparado: cmp.historico,
              Diferenca: round2(bestDiff),
            });
          }
        });

        referenceRows.forEach((ref, refIndex) => {
          if (usedReferences.has(refIndex)) return;
          rows.push({
            Chave_Conciliacao: group.Chave_Conciliacao,
            Status_Par: 'Referência sem par na origem comparada',
            Origem_Referencia: referenceOrigin,
            Arquivo_Referencia: ref.sourceFile || '',
            Linha_Referencia: ref.linhaOriginal,
            Data_Referencia: formatDateBR(ref.data),
            Valor_Referencia: round2(ref.valorLiquido),
            Historico_Referencia: ref.historico,
            Origem_Comparada: origin,
            Arquivo_Comparado: '',
            Linha_Comparada: '',
            Data_Comparada: '',
            Valor_Comparado: '',
            Historico_Comparado: '',
            Diferenca: round2(Math.abs(ref.valorLiquido || 0)),
          });
        });

        candidates.forEach((cmp, candidateIndex) => {
          if (usedCandidates.has(candidateIndex)) return;
          rows.push({
            Chave_Conciliacao: group.Chave_Conciliacao,
            Status_Par: 'Comparado sem par na referência',
            Origem_Referencia: referenceOrigin,
            Arquivo_Referencia: '',
            Linha_Referencia: '',
            Data_Referencia: '',
            Valor_Referencia: '',
            Historico_Referencia: '',
            Origem_Comparada: origin,
            Arquivo_Comparado: cmp.sourceFile || '',
            Linha_Comparada: cmp.linhaOriginal,
            Data_Comparada: formatDateBR(cmp.data),
            Valor_Comparado: round2(cmp.valorLiquido),
            Historico_Comparado: cmp.historico,
            Diferenca: round2(Math.abs(cmp.valorLiquido || 0)),
          });
        });
      });
    });
    return rows;
  }

  function buildReconciliationPairs(groups, mapping) {
    if (mapping.reconciliationScope === 'between_files' || mapping.reconciliationMatchMode === 'source_totals') {
      return buildReconciliationPairsBetweenFiles(groups, mapping.tolerance);
    }
    return buildReconciliationPairsWithinFile(groups, mapping.tolerance);
  }

  function groupAllByNF(items) {
    const map = new Map();
    items.forEach((item) => {
      if (!item.nf) return;
      const key = `${item.nf}|${item.nomeExtraido || 'NAO IDENTIFICADO'}`;
      if (!map.has(key)) {
        map.set(key, {
          Nota_Fiscal: item.nf,
          Nome_Extraido: item.nomeExtraido || 'NAO IDENTIFICADO',
          Qtde_Linhas: 0,
          Valor_Positivo: 0,
          Valor_Negativo: 0,
          Saldo: 0,
        });
      }
      const row = map.get(key);
      row.Qtde_Linhas += 1;
      if ((item.valorLiquido || 0) > 0) row.Valor_Positivo += item.valorLiquido || 0;
      if ((item.valorLiquido || 0) < 0) row.Valor_Negativo += item.valorLiquido || 0;
      row.Saldo += item.valorLiquido || 0;
    });
    return [...map.values()].map((row) => ({
      ...row,
      Valor_Positivo: round2(row.Valor_Positivo),
      Valor_Negativo: round2(row.Valor_Negativo),
      Saldo: round2(row.Saldo),
    })).sort((a, b) => Math.abs(b.Saldo || 0) - Math.abs(a.Saldo || 0));
  }

  function groupAllByExtractions(items, extractionDefs) {
    const map = new Map();
    (extractionDefs || []).forEach((def) => {
      items.forEach((item) => {
        const value = item.extracoes?.[def.key];
        if (!value) return;
        const key = `${def.key}|${value}`;
        if (!map.has(key)) {
          map.set(key, {
            Campo_Extraido: def.label,
            Valor_Extraido: value,
            Qtde_Linhas: 0,
            Valor_Positivo: 0,
            Valor_Negativo: 0,
            Saldo: 0,
          });
        }
        const row = map.get(key);
        row.Qtde_Linhas += 1;
        if ((item.valorLiquido || 0) > 0) row.Valor_Positivo += item.valorLiquido || 0;
        if ((item.valorLiquido || 0) < 0) row.Valor_Negativo += item.valorLiquido || 0;
        row.Saldo += item.valorLiquido || 0;
      });
    });
    return [...map.values()].map((row) => ({
      ...row,
      Valor_Positivo: round2(row.Valor_Positivo),
      Valor_Negativo: round2(row.Valor_Negativo),
      Saldo: round2(row.Saldo),
    })).sort((a, b) => {
      if (a.Campo_Extraido !== b.Campo_Extraido) return a.Campo_Extraido.localeCompare(b.Campo_Extraido);
      return Math.abs(b.Saldo || 0) - Math.abs(a.Saldo || 0);
    });
  }

  function analyze() {
    const mapping = getSelectedMapping();
    const periodMode = els.periodModeSelect.value;
    const selectedSources = selectedSourceDefsForAnalysis(mapping);
    if (!selectedSources.length) throw new Error('Nenhuma fonte selecionada para análise.');
    const baseItems = buildBaseItems(selectedSources, mapping, periodMode);
    if (!baseItems.length) throw new Error('Nenhuma linha válida encontrada na planilha. Confira as colunas selecionadas.');

    const years = getAvailableYears(baseItems);
    const items = baseItems.map((item) => ({ ...item, excluidaDuplicidade: false }));
    let duplicateBlocks = [];
    let removed = new Set();

    if (els.removeDuplicatesToggle.checked) {
      const duplicates = findDuplicateBlocksBySource(items, mapping.filterDefs);
      duplicateBlocks = duplicates.blocks;
      removed = duplicates.removed;
      items.forEach((item, idx) => { item.excluidaDuplicidade = removed.has(idx); });
    }

    const usedItems = items.filter((item) => !item.excluidaDuplicidade);
    state.baseItems = baseItems;

    if (mapping.analysisMode === 'reconciliation') {
      const reconciliationGroupsRaw = buildReconciliationGroups(usedItems, mapping);
      const reconciliationGroups = reconciliationGroupsRaw.map((group) => {
        const clean = { ...group };
        delete clean._linhas;
        delete clean._originTotals;
        return clean;
      });
      const reconciliationLines = buildReconciliationLines(reconciliationGroupsRaw);
      const reconciliationPairs = buildReconciliationPairs(reconciliationGroupsRaw, mapping);
      const nfSummaryAll = groupAllByNF(usedItems);
      const extractionSummaryAll = groupAllByExtractions(usedItems, mapping.extractionDefs);
      const totalSaldo = usedItems.reduce((acc, item) => acc + item.valorLiquido, 0);
      const divergentGroups = reconciliationGroups.filter((row) => row.Status_Conciliacao !== 'Conciliado' && row.Status_Conciliacao !== 'Conciliado - um lado');
      const reconciledGroups = reconciliationGroups.length - divergentGroups.length;
      const saldoDivergente = divergentGroups.reduce((acc, row) => acc + Math.abs(row.Diferenca_Absoluta ?? row.Saldo ?? 0), 0);

      state.analysis = {
        mode: 'reconciliation',
        detectedYears: years,
        periodMode,
        periodLabel: PERIOD_LABELS[periodMode] || periodMode,
        mapping,
        selectedSources,
        allItems: items,
        usedItems,
        reconciliationGroups,
        reconciliationLines,
        reconciliationPairs,
        nfSummaryAll,
        extractionSummaryAll,
        duplicateBlocks,
        totals: {
          originalRows: baseItems.length,
          usedRows: usedItems.length,
          duplicatesRemoved: removed.size,
          names: new Set(usedItems.map((item) => item.nomeExtraido).filter(Boolean)).size,
          invoices: new Set(usedItems.map((item) => item.nf).filter(Boolean)).size,
          extractedValues: countExtractedValues(usedItems, mapping.extractionDefs),
          groups: reconciliationGroups.length,
          reconciledGroups,
          divergentGroups: divergentGroups.length,
          saldo: round2(totalSaldo),
          saldoDivergente: round2(saldoDivergente),
        },
      };

      renderAnalysis();
      return;
    }

    if (years.length < 2) {
      throw new Error('A coluna de data selecionada não contém pelo menos dois anos diferentes para comparar. Para bases de um único ano, use o modo Conciliação.');
    }

    const currentYear = years[0];
    const previousYear = years[1];
    const comparison = groupByName(usedItems, currentYear, previousYear, mapping.filterDefs);
    const periodSummary = groupByPeriod(usedItems, currentYear, previousYear);
    const categorySummary = groupByCategory(usedItems, currentYear, previousYear);
    const filterSummary = groupByFilters(usedItems, currentYear, previousYear, mapping.filterDefs);
    const nfSummary = groupByNF(usedItems, currentYear, previousYear, mapping.filterDefs);
    const extractionSummary = groupByExtractions(usedItems, currentYear, previousYear, mapping.extractionDefs);
    const extractionNameSummary = groupByExtractionAndName(usedItems, currentYear, previousYear, mapping.extractionDefs);
    const totalCurrent = usedItems.filter((item) => item.ano === currentYear).reduce((acc, item) => acc + item.valorLiquido, 0);
    const totalPrevious = usedItems.filter((item) => item.ano === previousYear).reduce((acc, item) => acc + item.valorLiquido, 0);

    state.analysis = {
      mode: 'comparison',
      currentYear,
      previousYear,
      detectedYears: years,
      periodMode,
      periodLabel: PERIOD_LABELS[periodMode] || periodMode,
      mapping,
      selectedSources,
      allItems: items,
      usedItems,
      comparison,
      periodSummary,
      categorySummary,
      filterSummary,
      nfSummary,
      extractionSummary,
      extractionNameSummary,
      duplicateBlocks,
      totals: {
        originalRows: baseItems.length,
        usedRows: usedItems.length,
        duplicatesRemoved: removed.size,
        names: new Set(comparison.map((row) => row.Nome_Extraido)).size,
        invoices: new Set(usedItems.map((item) => item.nf).filter(Boolean)).size,
        extractedValues: countExtractedValues(usedItems, mapping.extractionDefs),
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

    if (analysis.mode === 'reconciliation') {
      els.metricComparison.textContent = 'Conciliação';
      els.metricPeriodMode.textContent = analysis.periodLabel;
      els.metricUsed.textContent = formatNumber(analysis.totals.usedRows);
      els.metricDuplicates.textContent = formatNumber(analysis.totals.duplicatesRemoved);
      els.metricNames.textContent = formatNumber(analysis.totals.groups);
      if (els.metricInvoices) els.metricInvoices.textContent = formatNumber(analysis.totals.invoices);
      if (els.metricExtractions) els.metricExtractions.textContent = formatNumber(analysis.totals.extractedValues);
      if (els.metricResultLabel) els.metricResultLabel.textContent = 'Saldo divergente';
      els.metricVariation.textContent = formatMoney(analysis.totals.saldoDivergente);

      els.summarySection.classList.remove('hidden');
      els.previewSection.classList.remove('hidden');
      els.exportBtn.disabled = false;
      els.resetBtn.disabled = false;
      els.filterControls.classList.add('hidden');
      els.filterControls.innerHTML = '';
      renderPreviewTable();

      const duplicateText = analysis.totals.duplicatesRemoved
        ? `${formatNumber(analysis.totals.duplicatesRemoved)} linhas removidas por duplicidade de bloco.`
        : 'Nenhum bloco duplicado removido.';
      const scopeText = analysis.mapping.reconciliationScope === 'between_files' ? 'entre arquivos/abas' : 'dentro de um arquivo/aba';
      setStatus(
        `Conciliação ${scopeText} concluída. ${formatNumber(analysis.totals.reconciledGroups)} grupos conciliados e ${formatNumber(analysis.totals.divergentGroups)} grupos pendentes. ${duplicateText}`,
        'success',
      );
      return;
    }

    els.metricComparison.textContent = `${analysis.currentYear} x ${analysis.previousYear}`;
    els.metricPeriodMode.textContent = analysis.periodLabel;
    els.metricUsed.textContent = formatNumber(analysis.totals.usedRows);
    els.metricDuplicates.textContent = formatNumber(analysis.totals.duplicatesRemoved);
    els.metricNames.textContent = formatNumber(analysis.totals.names);
    if (els.metricInvoices) els.metricInvoices.textContent = formatNumber(analysis.totals.invoices);
    if (els.metricExtractions) els.metricExtractions.textContent = formatNumber(analysis.totals.extractedValues);
    if (els.metricResultLabel) els.metricResultLabel.textContent = 'Variação total';
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

    if (analysis.mode === 'reconciliation') {
      const rows = analysis.reconciliationGroups
        .filter((row) => {
          if (!query) return true;
          return normalizeText(Object.values(row).join(' | ')).includes(query);
        })
        .slice(0, 25);

      const dynamicKeyHeaders = (analysis.mapping.reconciliationKeys || []).map((key) => key.label);
      const headers = [
        'Status_Conciliacao',
        'Observacao_Conciliacao',
        ...dynamicKeyHeaders,
        'Qtde_Linhas',
        'Qtde_Debitos',
        'Qtde_Creditos',
        'Valor_Positivo',
        'Valor_Negativo',
        'Saldo',
        'Diferenca_Absoluta',
        'Origens_Lados',
      ];

      els.previewTable.querySelector('thead').innerHTML = `<tr>${headers.map((h) => `<th>${escapeHTML(h)}</th>`).join('')}</tr>`;
      els.previewTable.querySelector('tbody').innerHTML = rows.map((row) => {
        return `<tr>${headers.map((header) => {
          const value = row[header];
          const isNumeric = /^Qtde_|^Valor_|Saldo|Diferenca/.test(header);
          let display = value;
          if (/^Valor_|Saldo|Diferenca/.test(header)) display = formatMoney(value);
          return `<td class="${isNumeric ? 'numeric' : ''}">${escapeHTML(display ?? '-')}</td>`;
        }).join('')}</tr>`;
      }).join('');
      return;
    }

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

  async function handleFiles(files) {
    const fileList = Array.from(files || []).filter(Boolean);
    if (!fileList.length) return;
    resetState(false);
    setStatus('Lendo arquivo(s)...', 'muted');

    state.workbooks = [];
    state.sourceDefs = [];
    state.rowsBySheet.clear();

    for (let fileIndex = 0; fileIndex < fileList.length; fileIndex += 1) {
      const file = fileList[fileIndex];
      const workbookName = file.name.replace(/\.[^.]+$/, '');
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
      state.workbooks.push({ fileName: file.name, workbookName, workbook });

      workbook.SheetNames.forEach((sheetName, sheetIndex) => {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
        const key = `${fileIndex}::${sheetIndex}::${workbookName}::${sheetName}`;
        const label = fileList.length > 1 ? `${file.name} / ${sheetName}` : sheetName;
        const headerRowIndex = findHeaderRow(rows);
        const columnDefs = buildColumnDefs(rows, headerRowIndex);
        const defaults = detectDefaultColumns(rows, headerRowIndex, columnDefs);
        const source = {
          key,
          fileIndex,
          sheetIndex,
          fileName: file.name,
          workbookName,
          sheetName,
          label,
          rows,
          headerRowIndex,
          columnDefs,
          defaults,
        };
        state.sourceDefs.push(source);
        state.rowsBySheet.set(key, rows);
      });
    }

    state.workbook = state.workbooks[0]?.workbook || null;
    state.workbookName = fileList.length === 1
      ? fileList[0].name.replace(/\.[^.]+$/, '')
      : `${fileList.length} arquivos importados`;

    els.sheetSelect.innerHTML = '';
    els.sourceSheetsSelect.innerHTML = '';
    state.sourceDefs.forEach((source) => {
      const option = document.createElement('option');
      option.value = source.key;
      option.textContent = source.label;
      els.sheetSelect.appendChild(option);

      const sourceOption = document.createElement('option');
      sourceOption.value = source.key;
      sourceOption.textContent = source.label;
      sourceOption.selected = true;
      els.sourceSheetsSelect.appendChild(sourceOption);
    });

    els.sheetSelect.disabled = state.sourceDefs.length < 2;
    els.resetBtn.disabled = false;
    if (!state.sourceDefs.length) throw new Error('Nenhum arquivo contém abas legíveis.');
    loadSheet(state.sourceDefs[0].key);
    refreshSourceSelectionDefaults();
    updateReconciliationScopeUI();
  }

  function toBaseExportRows(items, filterDefs, extractionDefs) {
    return items.map((item) => {
      const row = {
        Arquivo_Origem: item.sourceFile || '',
        Aba_Origem: item.sourceSheet || '',
        Linha_Original: item.linhaOriginal,
        Data: formatDateBR(item.data),
        Ano: item.ano,
        Mes: item.mes,
        Periodo: item.periodo,
        Origem_Lado: item.origemLado || '',
        Categoria_Tratada: item.categoriaTratada,
        Nome_Extraido: item.nomeExtraido,
        Nota_Fiscal: item.nf,
      };
      filterDefs.forEach((def) => { row[def.label] = item.filterValues[def.label] || ''; });
      (extractionDefs || []).forEach((def) => {
        const columnName = `Ext_${def.label}`.replace(/\s+/g, '_');
        if (columnName !== 'Ext_Nota_Fiscal') row[columnName] = item.extracoes?.[def.key] || '';
      });
      return {
        ...row,
        Historico: item.historico,
        Debito: round2(item.debito),
        Credito: round2(item.credito),
        Valor_Liquido: round2(item.valorLiquido),
        Metodo_Extracao: item.metodoExtracao,
        Excluida_Duplicidade: item.excluidaDuplicidade ? 'Sim' : 'Não',
      };
    });
  }

  function selectedColumnName(index) {
    if (index === null || index === undefined || index === '') return 'Não usado';
    const def = state.columnDefs.find((col) => col.index === Number(index));
    return def ? `${def.letter} - ${def.label}` : String(index);
  }

  function selectedSourceNames(analysis) {
    const sources = analysis?.selectedSources || [];
    if (!sources.length) return els.sheetSelect?.selectedOptions?.[0]?.textContent || state.workbookName || 'Fonte importada';
    return sources.map((source) => source.label).join('; ');
  }

  function buildSummaryRows(analysis) {
    const mapping = analysis.mapping;
    const filterNames = mapping.filterDefs.map((def) => `${def.letter} - ${def.label}`).join(', ') || 'Nenhuma';
    const extractionNames = mapping.extractionDefs.map((def) => def.label).join(', ') || 'Nenhuma';

    if (analysis.mode === 'reconciliation') {
      const keyNames = mapping.reconciliationKeys.map((key) => key.label).join(', ') || 'Nenhuma';
      const rows = [
        ['Conciliação por histórico e chaves selecionadas'],
        [],
        ['Arquivo(s)/fonte(s) analisados', selectedSourceNames(analysis)],
        ['Fonte usada para mapear colunas', els.sheetSelect?.selectedOptions?.[0]?.textContent || els.sheetSelect.value],
        ['Tipo de análise', 'Conciliação'],
        ['Escopo da conciliação', mapping.reconciliationScope === 'between_files' ? 'Entre arquivos/abas' : 'Dentro de um único arquivo/aba'],
        ['Visão de período', analysis.periodLabel],
        ['Anos detectados', analysis.detectedYears.join(', ') || 'Nenhum ano identificado'],
        ['Coluna de data', selectedColumnName(mapping.data)],
        ['Coluna de histórico', selectedColumnName(mapping.historico)],
        ['Coluna de débito', selectedColumnName(mapping.debito)],
        ['Coluna de crédito', selectedColumnName(mapping.credito)],
        ['Coluna de valor pronto', selectedColumnName(mapping.valor)],
        ['Coluna de origem/lado', selectedColumnName(mapping.sideColumn)],
        ['Colunas de filtros adicionais', filterNames],
        ['Campos extraídos do histórico', extractionNames],
        ['Chaves de conciliação', keyNames],
        ['Tolerância de valor', mapping.tolerance],
        ['Exige sinais opostos', mapping.reconciliationScope === 'between_files' ? 'Não se aplica ao modo entre arquivos' : (mapping.requireOppositeSigns ? 'Sim' : 'Não')],
        ['Linhas originais lidas', analysis.totals.originalRows],
        ['Linhas usadas na conciliação', analysis.totals.usedRows],
        ['Linhas excluídas como duplicação de bloco', analysis.totals.duplicatesRemoved],
        ['Grupos de conciliação', analysis.totals.groups],
        ['Grupos conciliados', analysis.totals.reconciledGroups],
        ['Grupos pendentes/divergentes', analysis.totals.divergentGroups],
        ['Notas fiscais identificadas', analysis.totals.invoices],
        ['Valores extraídos distintos', analysis.totals.extractedValues],
        ['Saldo total da base', analysis.totals.saldo],
        ['Saldo absoluto dos grupos pendentes', analysis.totals.saldoDivergente],
        [],
        ['Top 15 grupos pendentes por diferença'],
        ['Status', 'Observação', 'Chave_Conciliacao', 'Qtde_Linhas', 'Valor_Positivo', 'Valor_Negativo', 'Saldo', 'Diferenca_Absoluta'],
      ];

      analysis.reconciliationGroups
        .filter((row) => row.Status_Conciliacao !== 'Conciliado' && row.Status_Conciliacao !== 'Conciliado - um lado')
        .slice(0, 15)
        .forEach((row) => {
          rows.push([
            row.Status_Conciliacao,
            row.Observacao_Conciliacao || '',
            row.Chave_Conciliacao,
            row.Qtde_Linhas,
            row.Valor_Positivo,
            row.Valor_Negativo,
            row.Saldo,
            row.Diferenca_Absoluta,
          ]);
        });

      return rows;
    }

    const variationPct = analysis.totals.previous !== 0 ? analysis.totals.variation / analysis.totals.previous : null;
    const rows = [
      [`Comparativo ${analysis.currentYear} x ${analysis.previousYear} por nome extraído do histórico`],
      [],
      ['Arquivo/fonte analisado', selectedSourceNames(analysis)],
      ['Fonte usada para mapear colunas', els.sheetSelect?.selectedOptions?.[0]?.textContent || els.sheetSelect.value],
      ['Tipo de análise', 'Comparação'],
      ['Visão de período', analysis.periodLabel],
      ['Anos detectados', analysis.detectedYears.join(', ')],
      ['Anos comparados automaticamente', `${analysis.currentYear} x ${analysis.previousYear}`],
      ['Coluna de data', selectedColumnName(mapping.data)],
      ['Coluna de histórico', selectedColumnName(mapping.historico)],
      ['Coluna de débito', selectedColumnName(mapping.debito)],
      ['Coluna de crédito', selectedColumnName(mapping.credito)],
      ['Coluna de valor pronto', selectedColumnName(mapping.valor)],
      ['Colunas de filtros adicionais', filterNames],
      ['Campos extraídos do histórico', extractionNames],
      ['Linhas originais lidas', analysis.totals.originalRows],
      ['Linhas usadas no comparativo', analysis.totals.usedRows],
      ['Linhas excluídas como duplicação de bloco', analysis.totals.duplicatesRemoved],
      ['Regra aplicada para autônomos', 'AUTONOMO + INSS S/ PF = AUTONOMO'],
      ['Regra de nota fiscal', 'Extrai NF, NFE, NF-e, NFS-e e NOTA FISCAL do histórico, quando houver número identificável'],
      ['Critério do valor', mapping.valor === null ? 'Valor líquido = Débito - Crédito' : 'Valor líquido = coluna de valor pronto selecionada'],
      ['Nomes identificados', analysis.totals.names],
      ['Notas fiscais identificadas', analysis.totals.invoices],
      ['Valores extraídos distintos', analysis.totals.extractedValues],
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
      Object.keys(copy).forEach((key) => {
        if (key.startsWith('_')) delete copy[key];
      });
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
    const extractionCount = analysis.mapping.extractionDefs.length;
    const wb = XLSX.utils.book_new();
    const summarySheet = aoaSheet(buildSummaryRows(analysis), [54, 38, 18, 18, 18, 24]);
    XLSX.utils.book_append_sheet(wb, summarySheet, 'Resumo');

    if (analysis.mode === 'reconciliation') {
      const groupSheet = jsonSheet(hideInternalColumns(analysis.reconciliationGroups), [60, ...Array((analysis.mapping.reconciliationKeys || []).length).fill(30), 22, 12, 12, 12, 12, 18, 18, 18, 18, 18, 20, 20]);
      const lineSheet = jsonSheet(analysis.reconciliationLines, [60, 22, 16, 14, 12, 10, 18, 24, 26, 48, 16, 80, 14, 14, 16, 20]);
      const pairSheet = jsonSheet(analysis.reconciliationPairs, [60, 20, 14, 12, 16, 80, 14, 12, 16, 80, 16]);
      const nfAllSheet = jsonSheet(analysis.nfSummaryAll, [16, 48, 12, 18, 18, 18]);
      const extractionAllSheet = jsonSheet(analysis.extractionSummaryAll, [28, 36, 12, 18, 18, 18]);
      const baseSheet = jsonSheet(toBaseExportRows(analysis.usedItems, analysis.mapping.filterDefs, analysis.mapping.extractionDefs), [14, 12, 10, 8, 18, 24, 24, 48, 16, ...Array(filterCount).fill(30), ...Array(Math.max(extractionCount - 1, 0)).fill(24), 80, 14, 14, 16, 20]);
      const duplicatesSheet = jsonSheet(analysis.duplicateBlocks, [24, 24, 24, 24, 18]);

      XLSX.utils.book_append_sheet(wb, groupSheet, 'Conciliacao_Grupos');
      XLSX.utils.book_append_sheet(wb, lineSheet, 'Conciliacao_Linhas');
      XLSX.utils.book_append_sheet(wb, pairSheet, 'Conciliacao_Pares');
      XLSX.utils.book_append_sheet(wb, nfAllSheet, 'Resumo_NF');
      XLSX.utils.book_append_sheet(wb, extractionAllSheet, 'Resumo_Extracoes');
      XLSX.utils.book_append_sheet(wb, baseSheet, 'Base_Tratada');
      XLSX.utils.book_append_sheet(wb, duplicatesSheet, 'Duplicidades');
    } else {
      const comparisonSheet = jsonSheet(hideInternalColumns(analysis.comparison), [18, 48, ...Array(filterCount).fill(28), 12, 12, 16, 16, 16, 14, 16, 24]);
      const periodSheet = jsonSheet(hideInternalColumns(analysis.periodSummary), [18, 12, 12, 16, 16, 16, 14]);
      const categorySheet = jsonSheet(hideInternalColumns(analysis.categorySummary), [18, 28, 12, 12, 16, 16, 16, 14]);
      const filterSheet = jsonSheet(hideInternalColumns(analysis.filterSummary), [28, 36, 18, 12, 12, 16, 16, 16, 14]);
      const nfSheet = jsonSheet(hideInternalColumns(analysis.nfSummary), [18, 16, 48, 28, ...Array(filterCount).fill(30), 12, 12, 16, 16, 16, 14]);
      const extractionSheet = jsonSheet(hideInternalColumns(analysis.extractionSummary), [28, 36, 18, 12, 12, 16, 16, 16, 14]);
      const extractionNameSheet = jsonSheet(hideInternalColumns(analysis.extractionNameSummary), [28, 36, 18, 48, 12, 12, 16, 16, 16, 14]);
      const baseSheet = jsonSheet(toBaseExportRows(analysis.usedItems, analysis.mapping.filterDefs, analysis.mapping.extractionDefs), [14, 12, 10, 8, 18, 24, 24, 48, 16, ...Array(filterCount).fill(30), ...Array(Math.max(extractionCount - 1, 0)).fill(24), 80, 14, 14, 16, 20]);
      const duplicatesSheet = jsonSheet(analysis.duplicateBlocks, [24, 24, 24, 24, 18]);

      XLSX.utils.book_append_sheet(wb, comparisonSheet, 'Comparativo_Nome');
      XLSX.utils.book_append_sheet(wb, periodSheet, 'Resumo_Periodo');
      XLSX.utils.book_append_sheet(wb, categorySheet, 'Resumo_Categoria');
      XLSX.utils.book_append_sheet(wb, filterSheet, 'Resumo_Filtros');
      XLSX.utils.book_append_sheet(wb, nfSheet, 'Resumo_NF');
      XLSX.utils.book_append_sheet(wb, extractionSheet, 'Resumo_Extracoes');
      XLSX.utils.book_append_sheet(wb, extractionNameSheet, 'Extracao_Nome');
      XLSX.utils.book_append_sheet(wb, baseSheet, 'Base_Tratada');
      XLSX.utils.book_append_sheet(wb, duplicatesSheet, 'Duplicidades');
    }

    const safeName = (state.workbookName || 'analise')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9_-]+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '') || 'analise';

    if (analysis.mode === 'reconciliation') {
      XLSX.writeFile(wb, `${safeName}_conciliacao_${analysis.periodMode}.xlsx`);
    } else {
      XLSX.writeFile(wb, `${safeName}_comparativo_${analysis.currentYear}x${analysis.previousYear}_${analysis.periodMode}.xlsx`);
    }
  }

  function resetState(resetInput = true) {
    state.workbook = null;
    state.workbooks = [];
    state.workbookName = '';
    state.sourceDefs = [];
    state.activeSourceKey = '';
    state.rowsBySheet.clear();
    state.rawRows = [];
    state.headerRowIndex = -1;
    state.columnDefs = [];
    state.defaultColumns = null;
    state.baseItems = [];
    state.analysis = null;

    if (resetInput) els.fileInput.value = '';
    els.sheetSelect.innerHTML = '';
    if (els.sourceSheetsSelect) els.sourceSheetsSelect.innerHTML = '';
    if (els.analysisModeSelect) els.analysisModeSelect.value = 'comparison';
    if (els.reconciliationScopeSelect) els.reconciliationScopeSelect.value = 'single_file';
    if (els.reconciliationMatchModeSelect) els.reconciliationMatchModeSelect.value = 'balance';
    els.dateColumnSelect.innerHTML = '';
    els.historyColumnSelect.innerHTML = '';
    els.debitColumnSelect.innerHTML = '';
    els.creditColumnSelect.innerHTML = '';
    els.valueColumnSelect.innerHTML = '';
    els.filterColumnsSelect.innerHTML = '';
    if (els.reconciliationKeysSelect) els.reconciliationKeysSelect.innerHTML = '';
    if (els.sideColumnSelect) els.sideColumnSelect.innerHTML = '';
    if (els.toleranceInput) els.toleranceInput.value = '0.01';
    if (els.requireOppositeSignsToggle) els.requireOppositeSignsToggle.checked = true;
    if (els.extractionFieldsSelect) {
      Array.from(els.extractionFieldsSelect.options).forEach((option) => { option.selected = option.value === 'Nota_Fiscal'; });
    }
    if (els.customExtractionsInput) els.customExtractionsInput.value = '';
    els.sheetSelect.disabled = true;
    if (els.sourceSheetsBlock) els.sourceSheetsBlock.classList.add('hidden');
    enableMappingControls(false);
    els.analyzeBtn.disabled = true;
    els.exportBtn.disabled = true;
    els.resetBtn.disabled = true;
    els.summarySection.classList.add('hidden');
    els.previewSection.classList.add('hidden');
    els.filterControls.classList.add('hidden');
    els.filterControls.innerHTML = '';
    if (els.columnPreviewBlock) els.columnPreviewBlock.classList.add('hidden');
    if (els.columnPreviewTable) els.columnPreviewTable.querySelector('tbody').innerHTML = '';
    if (els.metricResultLabel) els.metricResultLabel.textContent = 'Resultado';
    updateAnalysisModeUI();
    els.searchInput.value = '';
    setStatus('Nenhum arquivo carregado.', 'muted');
  }

  ['change'].forEach((eventName) => {
    els.analysisModeSelect.addEventListener(eventName, () => { updateAnalysisModeUI(); els.exportBtn.disabled = true; });
    els.reconciliationScopeSelect.addEventListener(eventName, () => {
      if (els.reconciliationScopeSelect.value === 'between_files') {
        if (els.reconciliationMatchModeSelect) els.reconciliationMatchModeSelect.value = 'source_totals';
        if (els.requireOppositeSignsToggle) els.requireOppositeSignsToggle.checked = false;
      } else if (els.reconciliationMatchModeSelect) {
        els.reconciliationMatchModeSelect.value = 'balance';
        if (els.requireOppositeSignsToggle) els.requireOppositeSignsToggle.checked = true;
      }
      updateAnalysisModeUI();
      els.exportBtn.disabled = true;
    });
    if (els.reconciliationMatchModeSelect) els.reconciliationMatchModeSelect.addEventListener(eventName, () => { updateAnalysisModeUI(); els.exportBtn.disabled = true; });
    els.dateColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.historyColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.debitColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.creditColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.valueColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.filterColumnsSelect.addEventListener(eventName, () => { updateReconciliationKeyOptions(false); els.exportBtn.disabled = true; });
    els.extractionFieldsSelect.addEventListener(eventName, () => { updateReconciliationKeyOptions(false); els.exportBtn.disabled = true; });
    els.customExtractionsInput.addEventListener('input', () => { updateReconciliationKeyOptions(false); els.exportBtn.disabled = true; });
    els.reconciliationKeysSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.sourceSheetsSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.sideColumnSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.toleranceInput.addEventListener('input', () => { els.exportBtn.disabled = true; });
    els.requireOppositeSignsToggle.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
    els.periodModeSelect.addEventListener(eventName, () => { els.exportBtn.disabled = true; });
  });

  els.clearFilterColumnsBtn.addEventListener('click', () => {
    setMultiSelectValues(els.filterColumnsSelect, []);
    els.exportBtn.disabled = true;
  });

  els.selectDefaultExtractionsBtn.addEventListener('click', () => {
    Array.from(els.extractionFieldsSelect.options).forEach((option) => {
      option.selected = ['Nota_Fiscal', 'CNPJ_CPF', 'Competencia'].includes(option.value);
    });
    updateReconciliationKeyOptions(false);
    els.exportBtn.disabled = true;
  });

  els.clearExtractionsBtn.addEventListener('click', () => {
    Array.from(els.extractionFieldsSelect.options).forEach((option) => { option.selected = false; });
    els.customExtractionsInput.value = '';
    updateReconciliationKeyOptions(false);
    els.exportBtn.disabled = true;
  });

  els.selectDefaultReconciliationKeysBtn.addEventListener('click', () => {
    updateReconciliationKeyOptions(true);
    els.exportBtn.disabled = true;
  });

  els.clearReconciliationKeysBtn.addEventListener('click', () => {
    Array.from(els.reconciliationKeysSelect.options).forEach((option) => { option.selected = false; });
    els.exportBtn.disabled = true;
  });

  els.fileInput.addEventListener('change', (event) => {
    handleFiles(event.target.files).catch((error) => setStatus(`Erro ao ler arquivo: ${error.message}`, 'error'));
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
    handleFiles(event.dataTransfer.files).catch((error) => setStatus(`Erro ao ler arquivo: ${error.message}`, 'error'));
  });

  els.sheetSelect.addEventListener('change', () => loadSheet(els.sheetSelect.value));

  els.selectAllSourcesBtn.addEventListener('click', () => {
    Array.from(els.sourceSheetsSelect.options).forEach((option) => { option.selected = true; });
    els.exportBtn.disabled = true;
  });

  els.clearSourcesBtn.addEventListener('click', () => {
    Array.from(els.sourceSheetsSelect.options).forEach((option) => { option.selected = false; });
    els.exportBtn.disabled = true;
  });

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
