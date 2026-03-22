import { defaultProjectDefinitions } from './project-config.js';
import { defaultSubProjectDefinitions } from './subproject-config.js';
import {
  compileDefinitions,
  normalizeProjects,
  normalizeWords,
  normalizeDefinitions,
  matchesDefinition
} from './rule-engine.js';

const PROJECT_CONFIG_STORAGE_KEY = 'ticket-splitter.project-config.v1';
const SUBPROJECT_CONFIG_STORAGE_KEY = 'ticket-splitter.subproject-config.v1';

let projectDefinitions = loadDefinitions(PROJECT_CONFIG_STORAGE_KEY, defaultProjectDefinitions, {
  scoped: false,
  context: 'project'
});
let subProjectDefinitions = loadDefinitions(SUBPROJECT_CONFIG_STORAGE_KEY, defaultSubProjectDefinitions, {
  scoped: true,
  context: 'subproject'
});
let compiledProjectDefinitions = compileDefinitions(projectDefinitions, { scoped: false });
let compiledSubProjectDefinitions = compileDefinitions(subProjectDefinitions, { scoped: true });
let projectOrder = [...projectDefinitions.map((entry) => entry.name), 'Other'];

const navButtons = Array.from(document.querySelectorAll('[data-view]'));
const configNavButton = navButtons.find((button) => button.dataset.view === 'config') || null;
const importConfigTriggerButton = document.getElementById('import-config-trigger');
const splitterView = document.getElementById('splitter-view');
const configView = document.getElementById('config-view');
const fileInput = document.getElementById('report-input');
const hasHeaderCheckbox = document.getElementById('has-header');
const statusEl = document.getElementById('status');
const resultsSection = document.getElementById('results');
const summaryBody = document.getElementById('summary-body');
const zipButton = document.getElementById('download-zip');
const combinedButton = document.getElementById('download-combined');
const applyFiltersButton = document.getElementById('apply-filters');
const overallHoursEl = document.getElementById('overall-hours');
const configKindButtons = Array.from(document.querySelectorAll('[data-config-kind]'));
const configProjectFilterWrap = document.getElementById('config-project-filter-wrap');
const configProjectFilterSelect = document.getElementById('config-project-filter');
const configSearchInput = document.getElementById('config-search');
const configListEl = document.getElementById('config-list');
const addConfigGroupButton = document.getElementById('add-config-group');
const deleteConfigGroupButton = document.getElementById('delete-config-group');
const resetConfigButton = document.getElementById('reset-config');
const exportConfigButton = document.getElementById('export-config');
const importConfigInput = document.getElementById('import-config');
const configStatusEl = document.getElementById('config-status');
const configMainTitleEl = document.getElementById('config-main-title');
const configMainDescEl = document.getElementById('config-main-desc');
const configGroupCountEl = document.getElementById('config-group-count');
const configSaveToastEl = document.getElementById('config-save-toast');
const configGroupNameInput = document.getElementById('config-group-name');
const configScopePanel = document.getElementById('config-scope-panel');
const configScopeList = document.getElementById('config-scope-list');
const matcherBody = document.getElementById('matcher-body');
const addPatternButton = document.getElementById('add-pattern');
const runConfigTestButton = document.getElementById('run-config-test');
const configTestResultEl = document.getElementById('config-test-result');
const configTestPreviewEl = document.getElementById('config-test-preview');

let state = null;
let configSaveToastTimer = null;
let configUiState = {
  kind: getInitialConfigKind(),
  subprojectProjectFilter: getInitialSubprojectProjectFilter(),
  selectedIndexByKind: {
    project: 0,
    subproject: 0
  },
  search: ''
};

setupNavigation(getInitialView());
setupConfigPanel();
if (applyFiltersButton) {
  applyFiltersButton.disabled = true;
}
fileInput.addEventListener('change', handleFileSelection);
zipButton.addEventListener('click', downloadZipArchive);
combinedButton.addEventListener('click', downloadCombinedFile);
applyFiltersButton?.addEventListener('click', applyCurrentFilters);
configKindButtons.forEach((button) => {
  button.addEventListener('click', () => setConfigKind(button.dataset.configKind || 'project'));
});
configProjectFilterSelect?.addEventListener('change', handleProjectFilterChange);
configProjectFilterSelect?.addEventListener('input', handleProjectFilterChange);
configSearchInput?.addEventListener('input', (event) => {
  configUiState.search = normalizeSpaces(event.target.value || '').toLowerCase();
  renderConfigSidebar();
});
addConfigGroupButton.addEventListener('click', addConfigGroup);
deleteConfigGroupButton.addEventListener('click', deleteSelectedConfigGroup);
resetConfigButton.addEventListener('click', resetSelectedConfig);
exportConfigButton.addEventListener('click', exportConfig);
importConfigInput.addEventListener('change', importConfig);
importConfigTriggerButton?.addEventListener('click', () => {
  importConfigInput?.click();
});
configListEl.addEventListener('click', handleConfigListClick);
configGroupNameInput.addEventListener('input', handleGroupNameChange);
configScopeList?.addEventListener('change', handleScopeChange);
addPatternButton.addEventListener('click', addMatcherRow);
matcherBody.addEventListener('input', handleMatcherInput);
matcherBody.addEventListener('change', handleMatcherInput);
matcherBody.addEventListener('click', handleMatcherActionClick);
runConfigTestButton.addEventListener('click', runConfigTest);
summaryBody.addEventListener('click', (event) => {
  const target = event.target;
  if (target instanceof HTMLButtonElement && target.dataset.project) {
    downloadProject(target.dataset.project);
  }
});

function handleFileSelection(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  if (!isSupportedInputFile(file)) {
    setStatus('Please select a CSV or Excel file.', true);
    return;
  }

  setStatus('Reading file...');
  if (isSpreadsheetFile(file)) {
    readSpreadsheetFile(file);
    return;
  }

  readCsvFile(file);
}

function isSupportedInputFile(file) {
  return isCsvFile(file) || isSpreadsheetFile(file);
}

function isCsvFile(file) {
  const name = file.name.toLowerCase();
  const type = (file.type || '').toLowerCase();
  return name.endsWith('.csv') || type.includes('csv');
}

function isSpreadsheetFile(file) {
  const name = file.name.toLowerCase();
  const type = (file.type || '').toLowerCase();
  return (
    name.endsWith('.xls') ||
    name.endsWith('.xlsx') ||
    name.endsWith('.xlsm') ||
    name.endsWith('.xlsb') ||
    type === 'application/xlsx' ||
    type === 'application/vnd.ms-excel' ||
    type === 'application/vnd.ms-excel.sheet.binary.macroenabled.12' ||
    type === 'application/vnd.ms-excel.sheet.macroenabled.12' ||
    type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    type.includes('spreadsheetml')
  );
}

function readCsvFile(file) {
  Papa.parse(file, {
    skipEmptyLines: 'greedy',
    complete: (result) => processData(result.data),
    error: (error) => setStatus(`Failed to read file: ${error.message}`, true)
  });
}

function readSpreadsheetFile(file) {
  if (typeof XLSX === 'undefined') {
    setStatus('XLSX library failed to load. Check the CDN script tag.', true);
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const workbook = parseWorkbook(reader.result);
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        throw new Error('The Excel file has no sheets.');
      }
      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        raw: false,
        blankrows: false
      });
      processData(rows);
    } catch (error) {
      const message = formatSpreadsheetError(error);
      setStatus(`Failed to read file: ${message}`, true);
    }
  };
  reader.onerror = () => setStatus('Failed to read file.', true);
  reader.readAsArrayBuffer(file);
}

function parseWorkbook(arrayBuffer) {
  const parseAttempts = [
    () => XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' }),
    () => XLSX.read(arrayBufferToBinaryString(arrayBuffer), { type: 'binary' })
  ];

  let lastError = null;
  for (const attempt of parseAttempts) {
    try {
      return attempt();
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error('Unknown spreadsheet parsing error.');
}

function arrayBufferToBinaryString(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const chunkSize = 0x8000;
  const chunks = [];
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const slice = bytes.subarray(index, index + chunkSize);
    chunks.push(String.fromCharCode(...slice));
  }
  return chunks.join('');
}

function formatSpreadsheetError(error) {
  const message = error instanceof Error ? error.message : 'Unknown spreadsheet parsing error.';
  if (/encrypted|encryptioninfo|password/i.test(message)) {
    return 'This Excel file appears to be protected or saved in an unsupported format. Please open it in Excel or Google Sheets and save it again as a regular .xlsx or .csv file.';
  }
  return message;
}

function processData(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    setStatus('The file is empty or contains no rows.', true);
    return;
  }

  const normalized = rows
    .map((row) => (Array.isArray(row) ? row : Object.values(row)))
    .map((row) => {
      const cells = [];
      for (let index = 0; index < 4; index += 1) {
        cells.push(cleanCell(row[index]));
      }
      return cells;
    })
    .filter((row) => row.some((cell) => cell !== ''));

  if (!normalized.length) {
    setStatus('No data rows found.', true);
    return;
  }

  const defaultHeader = ['User', 'Comment', 'Duration', 'Date'];
  let header = defaultHeader.slice();
  let dataRows = normalized;

  if (hasHeaderCheckbox.checked) {
    const first = normalized[0];
    header = first.map((value, index) => (value ? value : defaultHeader[index]));
    dataRows = normalized.slice(1);
  }

  refreshDerivedState(header, dataRows);
}

function refreshDerivedState(header, dataRows) {
  const records = buildRecords(dataRows);
  const grouped = groupRecords(records);

  state = {
    header,
    rawRows: dataRows,
    records,
    grouped,
    total: dataRows.length
  };

  resultsSection.classList.remove('hidden');
  setStatus(`Processed rows: ${dataRows.length}`);
  updatePatternConfigAvailability();
  if (applyFiltersButton) {
    applyFiltersButton.disabled = !dataRows.length;
  }
  updateSummary();
}

function applyCurrentFilters() {
  if (!state?.rawRows?.length) {
    setStatus('Upload a file first, then apply filters.');
    return;
  }
  refreshDerivedState(state.header, state.rawRows);
  setStatus(`Filters applied. Processed rows: ${state.rawRows.length}`);
}

function buildRecords(dataRows) {
  const records = [];
  for (const row of dataRows) {
    const comment = normalizeSpaces(row[1]);
    const duration = parseDuration(row[2]);
    const project = classifyProject(comment);
    const subProject = classifySubProject(comment, project);
    records.push({ project, subProject, values: row, duration });
  }
  return records;
}

function groupRecords(records) {
  const grouped = new Map();
  for (const record of records) {
    if (!grouped.has(record.project)) {
      grouped.set(record.project, []);
    }
    grouped.get(record.project).push(record);
  }
  return grouped;
}

function getInitialView() {
  try {
    const url = new URL(window.location.href);
    const requestedView = url.searchParams.get('view') || url.hash.replace(/^#/, '');
    return requestedView === 'config' ? 'config' : 'splitter';
  } catch (error) {
    return 'splitter';
  }
}

function getInitialConfigKind() {
  try {
    const url = new URL(window.location.href);
    const requestedKind = String(url.searchParams.get('kind') || '').toLowerCase();
    return requestedKind === 'subproject' ? 'subproject' : 'project';
  } catch (error) {
    return 'project';
  }
}

function getInitialSubprojectProjectFilter() {
  return getRealProjectNames()[0] || '';
}

function setupNavigation(initialView = 'splitter') {
  updatePatternConfigAvailability();
  for (const button of navButtons) {
    button.addEventListener('click', () => {
      showView(button.dataset.view || 'splitter');
    });
  }
  showView(initialView);
}

function showView(viewName) {
  const blockedConfigView = viewName === 'config' && !hasLoadedFileData();
  const resolvedView = blockedConfigView ? 'splitter' : viewName;
  const isConfigView = resolvedView === 'config';
  if (blockedConfigView) {
    setStatus('Upload a file first, then Pattern Config will be available.');
  }
  splitterView?.classList.toggle('hidden', isConfigView);
  configView?.classList.toggle('hidden', !isConfigView);

  for (const button of navButtons) {
    button.classList.toggle('active', button.dataset.view === resolvedView);
  }
}

function hasLoadedFileData() {
  return Boolean(state?.rawRows?.length);
}

function updatePatternConfigAvailability() {
  if (!configNavButton) {
    return;
  }
  const enabled = hasLoadedFileData();
  configNavButton.disabled = !enabled;
  configNavButton.title = enabled ? '' : 'Upload CSV/XLSX first';
}

function setupConfigPanel() {
  renderConfigPanel();
  updateConfigStatus('Ready to edit rules.');
}

function renderConfigPanel() {
  renderConfigKindButtons();
  renderConfigSidebar();
  renderConfigEditor();
  renderConfigPreview();
}

function renderConfigKindButtons() {
  for (const button of configKindButtons) {
    button.classList.toggle('active', button.dataset.configKind === configUiState.kind);
  }
}

function getActiveDefinitions() {
  return configUiState.kind === 'project' ? projectDefinitions : subProjectDefinitions;
}

function getActiveStorageKey() {
  return configUiState.kind === 'project'
    ? PROJECT_CONFIG_STORAGE_KEY
    : SUBPROJECT_CONFIG_STORAGE_KEY;
}

function getActiveDefaults() {
  return configUiState.kind === 'project'
    ? defaultProjectDefinitions
    : defaultSubProjectDefinitions;
}

function getSelectedIndex() {
  return configUiState.selectedIndexByKind[configUiState.kind] ?? 0;
}

function setSelectedIndex(index) {
  configUiState.selectedIndexByKind[configUiState.kind] = index;
}

function getClampedSelectedIndex(definitions) {
  if (!Array.isArray(definitions) || !definitions.length) {
    return -1;
  }

  const rawIndex = Number(getSelectedIndex());
  if (!Number.isFinite(rawIndex)) {
    setSelectedIndex(0);
    return 0;
  }

  const clamped = Math.min(Math.max(rawIndex, 0), definitions.length - 1);
  if (clamped !== rawIndex) {
    setSelectedIndex(clamped);
  }
  return clamped;
}

function getSelectedDefinition() {
  const definitions = getActiveDefinitions();
  const index = getClampedSelectedIndex(definitions);
  if (index < 0) {
    return null;
  }
  return definitions[index] ?? null;
}

function getActiveProjectFilter() {
  if (configUiState.kind !== 'subproject') {
    return '';
  }
  return configUiState.subprojectProjectFilter || getRealProjectNames()[0] || '';
}

function setActiveProjectFilter(value) {
  configUiState.subprojectProjectFilter = String(value || '').trim();
}

function getRealProjectNames() {
  return projectDefinitions
    .map((entry) => entry.name)
    .filter((name) => name !== 'Incorrect codes');
}

function loadDefinitions(storageKey, defaults, options) {
  const stored = readStoredConfig(storageKey);
  if (!stored) {
    return cloneDefinitions(defaults);
  }

  try {
    const normalized = normalizeDefinitions(JSON.parse(stored), options);
    if (storageKey === SUBPROJECT_CONFIG_STORAGE_KEY) {
      const withoutLegacyTickets = normalized.filter((entry) => entry.name !== 'Tickets');
      return withoutLegacyTickets.length ? withoutLegacyTickets : cloneDefinitions(defaults);
    }
    return normalized;
  } catch (error) {
    console.warn(`Failed to load stored config for ${storageKey}, using defaults.`, error);
    return cloneDefinitions(defaults);
  }
}

function readStoredConfig(storageKey) {
  if (typeof localStorage === 'undefined') {
    return null;
  }
  try {
    return localStorage.getItem(storageKey);
  } catch (error) {
    return null;
  }
}

function writeStoredConfig(storageKey, value) {
  if (typeof localStorage === 'undefined') {
    return;
  }
  try {
    localStorage.setItem(storageKey, value);
  } catch (error) {
    console.warn(`Failed to persist config "${storageKey}".`, error);
  }
}

function cloneDefinitions(definitions) {
  return definitions.map((entry) => ({
    name: entry.name,
    projects: Array.isArray(entry.projects) ? [...entry.projects] : [],
    rules: Array.isArray(entry.rules)
      ? entry.rules.map((rule) => ({
          type: rule.type,
          value: rule.value,
          values: Array.isArray(rule.values) ? [...rule.values] : undefined
        }))
      : []
  }));
}

function normalizeConfigForSave(definitions, scoped) {
  const cleaned = definitions
    .map((entry) => {
      const name = String(entry.name ?? '').trim() || 'Untitled group';
      const projects = scoped ? normalizeProjects(entry.projects) : [];
      const rules = Array.isArray(entry.rules)
        ? entry.rules
            .map((rule) => normalizeRuleForSave(rule))
            .filter(Boolean)
        : [{ type: 'contains', value: '' }];
      const stableRules = rules.length ? rules : [{ type: 'contains', value: '' }];
      if (scoped && !projects.length) {
        const defaultProject = getActiveProjectFilter() || getRealProjectNames()[0] || '';
        return { name, projects: defaultProject ? [defaultProject] : [], rules: stableRules };
      }
      return { name, projects, rules: stableRules };
    });

  return normalizeDefinitions(cleaned, { scoped, context: scoped ? 'subproject' : 'project' });
}

function normalizeRuleForSave(rule) {
  if (!rule || typeof rule !== 'object') {
    return null;
  }

  const type = String(rule.type ?? 'contains').trim();
  if (type === 'contains') {
    const value = String(rule.value ?? '').trim();
    return { type, value };
  }

  if (type === 'containsWord') {
    const value = String(rule.value ?? '').trim();
    return { type, value };
  }

  if (type === 'containsAll') {
    const values = normalizeWords(rule.values ?? String(rule.value ?? '').split(','));
    return { type, values };
  }

  if (type === 'regex') {
    const value = String(rule.value ?? '').trim();
    return { type, value };
  }

  return null;
}

function renderConfigSidebar() {
  if (!configListEl) {
    return;
  }

  const definitions = getActiveDefinitions();
  const search = configUiState.search;
  const selectedIndex = getClampedSelectedIndex(definitions);
  renderConfigProjectFilter();
  const projectFilter = getActiveProjectFilter();
  const items = definitions
    .map((definition, index) => ({ definition, index }))
    .filter(({ definition }) => {
      if (configUiState.kind === 'subproject' && projectFilter) {
        const scope = Array.isArray(definition.projects) ? definition.projects : [];
        if (!scope.includes(projectFilter)) {
          return false;
        }
      }
      if (!search) {
        return true;
      }
      return definition.name.toLowerCase().includes(search);
    });

  configListEl.innerHTML = '';
  if (!items.length) {
    const empty = document.createElement('p');
    empty.className = 'config-empty';
    empty.textContent = configUiState.kind === 'subproject' && projectFilter
      ? `No sub project groups for "${projectFilter}".`
      : 'No groups match this search.';
    configListEl.appendChild(empty);
    return;
  }

  for (const { definition, index } of items) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'config-list-item';
    button.dataset.index = index.toString();
    button.classList.toggle('active', index === selectedIndex);

    const title = document.createElement('span');
    title.className = 'config-list-item-title';
    title.textContent = definition.name || 'Untitled group';

    const badge = document.createElement('span');
    badge.className = 'config-list-item-badge';
    badge.textContent = `${countMeaningfulMatchers(definition.rules)} patterns`;

    button.append(title, badge);
    configListEl.appendChild(button);
  }
}

function renderConfigProjectFilter() {
  if (!configProjectFilterWrap || !configProjectFilterSelect) {
    return;
  }

  const isSubproject = configUiState.kind === 'subproject';
  configProjectFilterWrap.classList.toggle('hidden', !isSubproject);
  if (!isSubproject) {
    return;
  }

  const projectNames = getRealProjectNames();
  const currentValue = getActiveProjectFilter();
  configProjectFilterSelect.innerHTML = '';

  if (!projectNames.length) {
    const option = document.createElement('option');
    option.value = '';
    option.textContent = 'No projects available';
    configProjectFilterSelect.appendChild(option);
    configProjectFilterSelect.disabled = true;
    return;
  }

  configProjectFilterSelect.disabled = false;
  for (const projectName of projectNames) {
    const option = document.createElement('option');
    option.value = projectName;
    option.textContent = projectName;
    configProjectFilterSelect.appendChild(option);
  }

  if (!projectNames.includes(currentValue)) {
    setActiveProjectFilter(projectNames[0]);
  }
  configProjectFilterSelect.value = getActiveProjectFilter();
}

function renderConfigEditor() {
  const definitions = getActiveDefinitions();
  const selected = getSelectedDefinition();
  const isProjectKind = configUiState.kind === 'project';

  if (configMainTitleEl) {
    configMainTitleEl.textContent = isProjectKind ? 'Project rules' : 'Sub project rules';
  }
  if (configMainDescEl) {
    configMainDescEl.textContent = isProjectKind
      ? 'Use pattern rows to decide which ticket goes to each project.'
      : 'Add a project scope, then define the sub project patterns below.';
  }
  if (configGroupCountEl) {
    configGroupCountEl.textContent = `${definitions.length} groups`;
  }

  if (!selected) {
    configGroupNameInput.value = '';
    if (configScopePanel) {
      configScopePanel.classList.add('hidden');
    }
    matcherBody.innerHTML = '';
    return;
  }

  configGroupNameInput.value = selected.name;
  configGroupNameInput.placeholder = isProjectKind ? 'Project name' : 'Sub project name';
  if (configSearchInput) {
    configSearchInput.placeholder = isProjectKind
      ? 'Type a project name'
      : 'Type a sub project name';
  }
  if (configScopePanel) {
    configScopePanel.classList.toggle('hidden', isProjectKind);
  }
  if (!isProjectKind) {
    renderScopeChecklist(selected.projects || []);
  }

  matcherBody.innerHTML = '';
  selected.rules.forEach((rule, index) => {
    matcherBody.appendChild(createMatcherRow(rule, index));
  });
}

function createMatcherRow(rule, index) {
  const row = document.createElement('tr');
  row.dataset.ruleIndex = index.toString();

  const typeCell = document.createElement('td');
  const typeSelect = document.createElement('select');
  typeSelect.dataset.action = 'type';
  for (const optionValue of ['contains', 'containsWord', 'containsAll', 'regex']) {
    const option = document.createElement('option');
    option.value = optionValue;
    option.textContent =
      optionValue === 'contains'
        ? 'Contains'
        : optionValue === 'containsWord'
          ? 'Contains word'
        : optionValue === 'containsAll'
          ? 'Contains all'
          : 'Regex';
    typeSelect.appendChild(option);
  }
  const resolvedType = ['contains', 'containsWord', 'containsAll', 'regex'].includes(rule.type)
    ? rule.type
    : 'contains';
  typeSelect.value = resolvedType;
  typeCell.appendChild(typeSelect);
  row.appendChild(typeCell);

  const valueCell = document.createElement('td');
  const input = document.createElement('input');
  input.type = 'text';
  input.dataset.action = 'value';
  input.value = formatRuleValue(rule);
  input.placeholder = getRulePlaceholder(rule.type);
  valueCell.appendChild(input);
  row.appendChild(valueCell);

  const actionsCell = document.createElement('td');
  const deleteButton = document.createElement('button');
  deleteButton.type = 'button';
  deleteButton.dataset.action = 'delete-rule';
  deleteButton.textContent = 'Remove';
  deleteButton.disabled = getSelectedDefinition()?.rules.length <= 1;
  actionsCell.appendChild(deleteButton);
  row.appendChild(actionsCell);

  return row;
}

function getRulePlaceholder(type) {
  if (type === 'containsAll') {
    return 'Word 1, Word 2';
  }
  if (type === 'containsWord') {
    return 'Exact word or ticket code';
  }
  if (type === 'regex') {
    return 'Advanced pattern';
  }
  return 'Type a phrase';
}

function formatRuleValue(rule) {
  if (rule.type === 'containsAll') {
    return Array.isArray(rule.values) ? rule.values.join(', ') : String(rule.value ?? '');
  }
  return String(rule.value ?? '');
}

function addConfigGroup() {
  const definitions = getActiveDefinitions();
  const isProjectKind = configUiState.kind === 'project';
  const newGroup = isProjectKind
    ? {
        name: 'New Project',
        rules: [{ type: 'contains', value: '' }]
      }
    : {
        name: 'New Sub Project',
        projects: [getActiveProjectFilter() || getRealProjectNames()[0] || ''],
        rules: [{ type: 'contains', value: '' }]
  };
  definitions.push(newGroup);
  setSelectedIndex(definitions.length - 1);
  renderConfigPanel();
  persistConfig({
    refreshSummary: true,
    refreshPreview: true,
    statusMessage: 'Group added. Changes saved.'
  });
}

function deleteSelectedConfigGroup() {
  const definitions = getActiveDefinitions();
  if (definitions.length <= 1) {
    updateConfigStatus('Keep at least one group.', true);
    return;
  }

  const index = getSelectedIndex();
  definitions.splice(index, 1);
  setSelectedIndex(Math.max(0, index - 1));
  renderConfigPanel();
  persistConfig({
    refreshSummary: true,
    refreshPreview: true,
    statusMessage: 'Group removed. Changes saved.'
  });
}

function handleConfigListClick(event) {
  const button = event.target.closest('.config-list-item');
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }
  const index = Number(button.dataset.index);
  if (!Number.isFinite(index)) {
    return;
  }
  setSelectedIndex(index);
  renderConfigPanel();
}

function setConfigKind(kind) {
  const nextKind = kind === 'subproject' ? 'subproject' : 'project';
  if (configUiState.kind === nextKind) {
    return;
  }
  configUiState.kind = nextKind;
  if (nextKind === 'subproject' && !getActiveProjectFilter()) {
    setActiveProjectFilter(getInitialSubprojectProjectFilter());
  }
  renderConfigPanel();
}

function handleGroupNameChange(event) {
  const selected = getSelectedDefinition();
  if (!selected) {
    return;
  }
  selected.name = String(event.target.value ?? '').trim();
  renderConfigSidebar();
  persistConfig({ refreshSummary: true, refreshPreview: true });
}

function handleScopeChange(event) {
  const selected = getSelectedDefinition();
  if (!selected || configUiState.kind !== 'subproject') {
    return;
  }
  const target = event.target;
  if (!(target instanceof HTMLInputElement) || target.type !== 'checkbox') {
    return;
  }

  const projectName = String(target.value ?? '').trim();
  const nextProjects = new Set(selected.projects || []);
  if (target.checked) {
    nextProjects.add(projectName);
  } else {
    nextProjects.delete(projectName);
  }
  selected.projects = normalizeProjects(Array.from(nextProjects));
  renderConfigSidebar();
  persistConfig({ refreshSummary: true, refreshPreview: true });
}

function handleProjectFilterChange(event) {
  if (configUiState.kind !== 'subproject') {
    return;
  }

  const nextValue = String(event.target.value || '').trim();
  setActiveProjectFilter(nextValue);
  const filtered = getActiveDefinitions()
    .map((definition, index) => ({ definition, index }))
    .filter(({ definition }) => (definition.projects || []).includes(nextValue));
  if (filtered.length) {
    setSelectedIndex(filtered[0].index);
  }
  renderConfigSidebar();
  renderConfigEditor();
  renderConfigPreview();
}

function addMatcherRow() {
  const selected = getSelectedDefinition();
  if (!selected) {
    return;
  }
  selected.rules.push({ type: 'contains', value: '' });
  renderConfigEditor();
  renderConfigSidebar();
  persistConfig({ refreshSummary: true, refreshPreview: true });
}

function handleMatcherInput(event) {
  const target = event.target;
  const row = target.closest('tr[data-rule-index]');
  if (!(row instanceof HTMLTableRowElement)) {
    return;
  }
  const index = Number(row.dataset.ruleIndex);
  const selected = getSelectedDefinition();
  if (!selected || !Number.isFinite(index) || !selected.rules[index]) {
    return;
  }

  if (target instanceof HTMLSelectElement && target.dataset.action === 'type') {
    const nextType = target.value;
    const currentValue = formatRuleValue(selected.rules[index]);
    selected.rules[index] = parseRuleFromValue(nextType, currentValue);
    renderConfigEditor();
    renderConfigSidebar();
    persistConfig({ refreshSummary: true, refreshPreview: true });
    return;
  }

  if (target instanceof HTMLInputElement && target.dataset.action === 'value') {
    const type = selected.rules[index].type;
    selected.rules[index] = parseRuleFromValue(type, target.value);
    persistConfig({ refreshSummary: true, refreshPreview: true });
  }
}

function handleMatcherActionClick(event) {
  const button = event.target.closest('button[data-action="delete-rule"]');
  if (!(button instanceof HTMLButtonElement)) {
    return;
  }
  const row = button.closest('tr[data-rule-index]');
  if (!(row instanceof HTMLTableRowElement)) {
    return;
  }
  const index = Number(row.dataset.ruleIndex);
  const selected = getSelectedDefinition();
  if (!selected || !Number.isFinite(index) || selected.rules.length <= 1) {
    return;
  }

  selected.rules.splice(index, 1);
  renderConfigEditor();
  renderConfigSidebar();
  persistConfig({ refreshSummary: true, refreshPreview: true });
}

function parseRuleFromValue(type, value) {
  if (type === 'containsAll') {
    const values = normalizeWords(String(value).split(','));
    return { type, values };
  }
  return { type, value: String(value).trim() };
}

function renderScopeChecklist(selectedProjects) {
  if (!configScopeList) {
    return;
  }

  const availableProjects = getRealProjectNames();
  configScopeList.innerHTML = '';

  if (!availableProjects.length) {
    const empty = document.createElement('p');
    empty.className = 'config-empty';
    empty.textContent = 'Create project rules first.';
    configScopeList.appendChild(empty);
    return;
  }

  for (const projectName of availableProjects) {
    const label = document.createElement('label');
    label.className = 'scope-option';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = projectName;
    checkbox.checked = selectedProjects.includes(projectName);

    const text = document.createElement('span');
    text.textContent = projectName;

    label.append(checkbox, text);
    configScopeList.appendChild(label);
  }
}

function persistConfig({ refreshSummary = true, refreshPreview = true, statusMessage = '' } = {}) {
  try {
    const nextProjectDefinitions = normalizeConfigForSave(projectDefinitions, false);
    const nextSubProjectDefinitions = normalizeConfigForSave(subProjectDefinitions, true);
    projectDefinitions = nextProjectDefinitions;
    subProjectDefinitions = nextSubProjectDefinitions;
    compiledProjectDefinitions = compileDefinitions(projectDefinitions, { scoped: false });
    compiledSubProjectDefinitions = compileDefinitions(subProjectDefinitions, { scoped: true });
    projectOrder = [...projectDefinitions.map((entry) => entry.name), 'Other'];
    writeStoredConfig(PROJECT_CONFIG_STORAGE_KEY, JSON.stringify(projectDefinitions));
    writeStoredConfig(SUBPROJECT_CONFIG_STORAGE_KEY, JSON.stringify(subProjectDefinitions));
    if (refreshSummary && state?.rawRows?.length) {
      refreshDerivedState(state.header, state.rawRows);
    }
    if (refreshPreview) {
      renderConfigPreview();
    }
    const resolvedStatusMessage = statusMessage || 'Changes saved.';
    showConfigSaveToast(resolvedStatusMessage);
    updateConfigStatus(resolvedStatusMessage);
  } catch (error) {
    updateConfigStatus(`Failed to save config: ${error.message}`, true);
  }
}

function resetSelectedConfig() {
  const defaults = cloneDefinitions(getActiveDefaults());
  if (configUiState.kind === 'project') {
    projectDefinitions = defaults;
    compiledProjectDefinitions = compileDefinitions(projectDefinitions, { scoped: false });
    projectOrder = [...projectDefinitions.map((entry) => entry.name), 'Other'];
    writeStoredConfig(PROJECT_CONFIG_STORAGE_KEY, JSON.stringify(projectDefinitions));
  } else {
    subProjectDefinitions = defaults;
    compiledSubProjectDefinitions = compileDefinitions(subProjectDefinitions, { scoped: true });
    writeStoredConfig(SUBPROJECT_CONFIG_STORAGE_KEY, JSON.stringify(subProjectDefinitions));
  }
  setSelectedIndex(0);
  renderConfigPanel();
  if (state?.rawRows?.length) {
    refreshDerivedState(state.header, state.rawRows);
  }
  renderConfigPreview();
  showConfigSaveToast('Defaults restored. Changes saved.');
  updateConfigStatus('Defaults restored. Changes saved.');
}

function exportConfig() {
  const payload = {
    projectDefinitions,
    subProjectDefinitions
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: 'application/json'
  });
  triggerDownload(blob, 'ticket-splitter-config.json');
  updateConfigStatus('Exported JSON backup.');
}

function importConfig(event) {
  const file = event.target.files?.[0];
  event.target.value = '';
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || '{}'));
      if (Array.isArray(parsed.projectDefinitions)) {
        projectDefinitions = normalizeDefinitions(parsed.projectDefinitions, { scoped: false, context: 'project' });
        compiledProjectDefinitions = compileDefinitions(projectDefinitions, { scoped: false });
        projectOrder = [...projectDefinitions.map((entry) => entry.name), 'Other'];
      }
      if (Array.isArray(parsed.subProjectDefinitions)) {
        subProjectDefinitions = normalizeDefinitions(parsed.subProjectDefinitions, { scoped: true, context: 'subproject' });
        compiledSubProjectDefinitions = compileDefinitions(subProjectDefinitions, { scoped: true });
      }
      writeStoredConfig(PROJECT_CONFIG_STORAGE_KEY, JSON.stringify(projectDefinitions));
      writeStoredConfig(SUBPROJECT_CONFIG_STORAGE_KEY, JSON.stringify(subProjectDefinitions));
      setSelectedIndex(0);
      renderConfigPanel();
      if (state?.rawRows?.length) {
        refreshDerivedState(state.header, state.rawRows);
      }
      showConfigSaveToast(`Imported "${file.name}". Changes saved.`);
      updateConfigStatus(`Imported "${file.name}". Changes saved.`);
    } catch (error) {
      updateConfigStatus(`Failed to import config: ${error.message}`, true);
    }
  };
  reader.onerror = () => updateConfigStatus(`Failed to read "${file.name}".`, true);
  reader.readAsText(file);
}

function runConfigTest() {
  renderConfigPreview();
}

function renderConfigPreview() {
  if (!configTestPreviewEl) {
    return;
  }

  const selected = getPreviewDefinition();
  if (!selected) {
    configTestPreviewEl.innerHTML = '';
    if (configTestResultEl) {
      configTestResultEl.textContent = 'Select a group to preview a row.';
    }
    return;
  }

  if (!hasMeaningfulMatchers(selected.rules)) {
    configTestPreviewEl.innerHTML = '';
    if (configTestResultEl) {
      configTestResultEl.textContent = 'Add at least one pattern to preview this group.';
    }
    const empty = document.createElement('p');
    empty.className = 'config-empty';
    empty.textContent = 'Add at least one pattern to preview this group.';
    configTestPreviewEl.appendChild(empty);
    return;
  }

  const previewState = getPreviewState(selected);
  const table = document.createElement('table');
  table.className = 'preview-table';

  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
      <th>User</th>
      <th>Comment</th>
      <th>Duration</th>
      <th>Date</th>
      <th>Project</th>
      <th>Sub Project</th>
    </tr>
  `;
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  const previewRows = Array.isArray(previewState.rows) ? previewState.rows : [];
  if (previewRows.length) {
    for (const previewRow of previewRows) {
      const row = document.createElement('tr');
      row.className = previewRow.previewMatchWinner === false ? 'preview-row no-match' : 'preview-row match';
      const values = Array.isArray(previewRow.values) ? previewRow.values : [];
      row.appendChild(createPreviewCell(values[0] || '—'));
      row.appendChild(createPreviewCell(values[1] || '—'));
      row.appendChild(createPreviewCell(values[2] || '—'));
      row.appendChild(createPreviewCell(values[3] || '—'));
      row.appendChild(createPreviewCell(previewRow.project || '—'));
      row.appendChild(createPreviewCell(previewRow.subProject || '—'));
      tbody.appendChild(row);
    }
    if (configTestResultEl) {
      configTestResultEl.textContent = previewState.message || `Preview for "${selected.name}".`;
    }
  } else {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 6;
    cell.textContent = `Unable to build a preview row for "${selected.name}".`;
    row.appendChild(cell);
    tbody.appendChild(row);
    if (configTestResultEl) {
      configTestResultEl.textContent = previewState.message || `Unable to build a preview row for "${selected.name}".`;
    }
  }

  table.appendChild(tbody);

  configTestPreviewEl.innerHTML = '';
  configTestPreviewEl.appendChild(table);
}

function getPreviewDefinition() {
  const selected = getSelectedDefinition();
  if (selected) {
    return selected;
  }

  const activeButton = configListEl?.querySelector('.config-list-item.active');
  if (activeButton instanceof HTMLButtonElement) {
    const index = Number(activeButton.dataset.index);
    const definitions = getActiveDefinitions();
    if (Number.isFinite(index) && definitions[index]) {
      return definitions[index];
    }
  }

  const definitions = getActiveDefinitions();
  return definitions[0] ?? null;
}

function getPreviewState(selected) {
  const uploadedPreview = getPreviewStateFromUploadedFile(selected);
  if (uploadedPreview) {
    return uploadedPreview;
  }

  const sampleRecord = buildSamplePreviewRecord(selected);
  const hasUploadedFile = Boolean(state?.rawRows?.length);
  return {
    source: 'sample',
    rows: sampleRecord ? [sampleRecord] : [],
    message: hasUploadedFile
      ? `No rows in the current file match "${selected.name}". Showing a sample row from the current pattern.`
      : `Upload a file to see real matches for "${selected.name}". Showing a sample row from the current pattern.`
  };
}

function getPreviewStateFromUploadedFile(selected) {
  if (!selected || !state?.rawRows?.length) {
    return null;
  }

  const compiledSelected = compileSelectedPreviewDefinition(selected);
  if (!compiledSelected) {
    return null;
  }

  const matchedRows = [];
  const winningRows = [];
  let patternMatches = 0;
  let winningMatches = 0;

  for (const row of state.rawRows) {
    const values = Array.isArray(row) ? row : [];
    const comment = normalizeSpaces(values[1]);
    const project = classifyProject(comment);
    if (!isPreviewScopeMatch(compiledSelected, project)) {
      continue;
    }
    if (!matchesDefinition(compiledSelected, comment)) {
      continue;
    }

    patternMatches += 1;
    const subProject = classifySubProject(comment, project);
    const isWinner = isWinningPreviewMatch(selected, project, subProject);
    const previewRecord = {
      values,
      project,
      subProject,
      previewMatchWinner: isWinner
    };
    matchedRows.push(previewRecord);

    if (isWinner) {
      winningMatches += 1;
      winningRows.push(previewRecord);
    }
  }

  if (matchedRows.length) {
    const label = configUiState.kind === 'project' ? 'project' : 'sub project';
    const rowsToShow = winningRows.length ? matchedRows : matchedRows;
    return {
      source: winningRows.length ? 'uploaded' : 'uploaded_non_winner',
      rows: rowsToShow,
      message: `${patternMatches} row(s) matched in file; ${winningMatches} assigned to this ${label}.`
    };
  }

  return null;
}

function compileSelectedPreviewDefinition(selected) {
  if (!selected) {
    return null;
  }

  try {
    const scoped = configUiState.kind === 'subproject';
    const normalized = normalizeDefinitions([selected], { scoped, context: 'preview' });
    const compiled = compileDefinitions(normalized, { scoped });
    return compiled[0] || null;
  } catch (error) {
    return null;
  }
}

function isPreviewScopeMatch(definition, project) {
  if (!definition) {
    return false;
  }
  if (configUiState.kind !== 'subproject') {
    return true;
  }
  if (!Array.isArray(definition.projects) || !definition.projects.length) {
    return true;
  }
  return definition.projects.includes(project);
}

function isWinningPreviewMatch(selected, project, subProject) {
  if (!selected) {
    return false;
  }
  if (configUiState.kind === 'project') {
    return project === selected.name;
  }
  return subProject === selected.name;
}

function buildSamplePreviewRecord(selected) {
  if (!selected) {
    return null;
  }

  const previewComment = buildPreviewComment(selected.rules);
  if (!previewComment) {
    return null;
  }

  const previewProject = getPreviewProject(selected, previewComment);
  const previewSubProject = getPreviewSubProject(selected, previewComment, previewProject);

  return {
    project: previewProject || '—',
    subProject: previewSubProject || '—',
    values: ['Preview', previewComment, '—', '—']
  };
}

function buildPreviewComment(rules) {
  if (!Array.isArray(rules)) {
    return '';
  }

  for (const rule of rules) {
    if (!rule || typeof rule !== 'object') {
      continue;
    }

    if (rule.type === 'containsAll') {
      const values = normalizeWords(rule.values);
      if (values.length) {
        return values.join(' ');
      }
      continue;
    }

    const rawValue = String(rule.value ?? '').trim();
    if (!rawValue) {
      continue;
    }

    if (rule.type === 'regex') {
      const simplified = rawValue.replace(/\\b/g, '').replace(/\\+/g, '').trim();
      return simplified || rawValue;
    }

    return rawValue;
  }

  return '';
}

function getPreviewProject(selected, previewComment) {
  if (configUiState.kind === 'project') {
    return selected.name || classifyProject(previewComment);
  }

  const scopedProjects = normalizeProjects(selected.projects);
  return scopedProjects[0] || getActiveProjectFilter() || classifyProject(previewComment) || 'Other';
}

function getPreviewSubProject(selected, previewComment, previewProject) {
  if (configUiState.kind === 'subproject') {
    return selected.name || classifySubProject(previewComment, previewProject);
  }
  return classifySubProject(previewComment, previewProject);
}

function hasMeaningfulMatchers(rules) {
  if (!Array.isArray(rules) || !rules.length) {
    return false;
  }

  return rules.some((rule) => {
    if (!rule || typeof rule !== 'object') {
      return false;
    }
    if (rule.type === 'containsAll') {
      return Array.isArray(rule.values) && rule.values.some((value) => String(value ?? '').trim());
    }
    return Boolean(String(rule.value ?? '').trim());
  });
}

function countMeaningfulMatchers(rules) {
  if (!Array.isArray(rules) || !rules.length) {
    return 0;
  }

  return rules.reduce((count, rule) => {
    if (!rule || typeof rule !== 'object') {
      return count;
    }
    if (rule.type === 'containsAll') {
      const values = normalizeWords(rule.values);
      return count + (values.length ? 1 : 0);
    }
    return count + (String(rule.value ?? '').trim() ? 1 : 0);
  }, 0);
}

function createPreviewCell(value) {
  const cell = document.createElement('td');
  cell.textContent = value;
  return cell;
}

function updateConfigStatus(message, isError = false) {
  if (!configStatusEl) {
    return;
  }
  configStatusEl.textContent = message;
  configStatusEl.classList.toggle('error', Boolean(isError));
  configStatusEl.classList.toggle('saved', !isError && /saved/i.test(String(message)));
}

function showConfigSaveToast(message) {
  if (!configSaveToastEl) {
    return;
  }

  configSaveToastEl.textContent = message;
  configSaveToastEl.classList.add('visible');
  clearTimeout(configSaveToastTimer);
  configSaveToastTimer = setTimeout(() => {
    configSaveToastEl.classList.remove('visible');
  }, 1800);
}

function classifyProject(comment) {
  for (const definition of compiledProjectDefinitions) {
    if (matchesDefinition(definition, comment)) {
      return definition.name;
    }
  }
  return 'Other';
}

function classifySubProject(comment, project) {
  for (const definition of compiledSubProjectDefinitions) {
    if (definition.projects.length && !definition.projects.includes(project)) {
      continue;
    }
    if (matchesDefinition(definition, comment)) {
      return definition.name;
    }
  }
  return '';
}

function updateSummary() {
  if (!state) {
    return;
  }

  summaryBody.innerHTML = '';
  const rows = [];
  let hoursAccumulator = 0;

  for (const project of projectOrder) {
    const entries = state.grouped.get(project) ?? [];
    if (!entries.length) {
      continue;
    }
    const totalHours = entries.reduce((sum, entry) => sum + entry.duration, 0);
    hoursAccumulator += totalHours;
    rows.push(createSummaryRow(project, entries.length, totalHours));
  }

  if (!rows.length) {
    const emptyRow = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 4;
    cell.textContent = 'No rows to display.';
    emptyRow.appendChild(cell);
    rows.push(emptyRow);
  }

  summaryBody.append(...rows);

  const hasAnyData = state.total > 0;
  zipButton.disabled =
    !hasAnyData ||
    typeof JSZip === 'undefined' ||
    typeof XLSX === 'undefined';
  combinedButton.disabled =
    typeof XLSX === 'undefined' ||
    !getRealProjectNames().some((project) => (state.grouped.get(project) ?? []).length > 0);
  overallHoursEl.textContent = hasAnyData
    ? `Overall hours: ${formatHours(hoursAccumulator)}`
    : '';
}

function createSummaryRow(project, count, hours) {
  const row = document.createElement('tr');

  const projectCell = document.createElement('td');
  projectCell.textContent = project;
  row.appendChild(projectCell);

  const countCell = document.createElement('td');
  countCell.textContent = count.toString();
  row.appendChild(countCell);

  const hoursCell = document.createElement('td');
  hoursCell.textContent = formatHours(hours);
  row.appendChild(hoursCell);

  const actionsCell = document.createElement('td');
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = 'XLSX';
  button.dataset.project = project;
  actionsCell.appendChild(button);
  row.appendChild(actionsCell);

  return row;
}

function downloadProject(project) {
  if (!state) {
    return;
  }
  if (typeof XLSX === 'undefined') {
    setStatus('XLSX library failed to load. Check the CDN script tag.', true);
    return;
  }
  const entries = state.grouped.get(project);
  if (!entries?.length) {
    setStatus(`No rows found for project "${project}".`, true);
    return;
  }
  const workbook = createWorkbookFromRecords(state.header, entries, {
    includeProject: false,
    includeSubProject: true,
    sheetName: project
  });
  const blob = workbookToBlob(workbook);
  triggerDownload(blob, `${toSafeFilename(project) || 'project'}.xlsx`);
  setStatus(`Workbook for project "${project}" ready for download.`);
}

function downloadCombinedFile() {
  if (!state) {
    return;
  }
  if (typeof XLSX === 'undefined') {
    setStatus('XLSX library failed to load. Check the CDN script tag.', true);
    return;
  }
  const combined = [];
  for (const project of getRealProjectNames()) {
    const entries = state.grouped.get(project);
    if (entries?.length) {
      combined.push(...entries);
    }
  }
  if (!combined.length) {
    setStatus('No rows available for the combined report.', true);
    return;
  }
  const workbook = createWorkbookFromRecords(state.header, combined, {
    includeProject: true,
    includeSubProject: true,
    sheetName: 'Projects'
  });
  const blob = workbookToBlob(workbook);
  triggerDownload(blob, 'all_projects.xlsx');
  setStatus('Combined workbook ready for download.');
}

function downloadZipArchive() {
  if (!state) {
    return;
  }
  if (typeof JSZip === 'undefined') {
    setStatus('JSZip failed to load. Check the CDN script tag.', true);
    return;
  }
  if (typeof XLSX === 'undefined') {
    setStatus('XLSX library failed to load. Check the CDN script tag.', true);
    return;
  }

  const zip = new JSZip();
  let filesAdded = 0;

  for (const [project, entries] of state.grouped.entries()) {
    if (!entries.length) {
      continue;
    }
    const workbook = createWorkbookFromRecords(state.header, entries, {
      includeProject: false,
      includeSubProject: true,
      sheetName: project
    });
    const buffer = workbookToArrayBuffer(workbook);
    zip.file(`${toSafeFilename(project) || 'project'}.xlsx`, buffer, { binary: true });
    filesAdded += 1;
  }

  if (!filesAdded) {
    setStatus('No data to export.', true);
    return;
  }

  zip.generateAsync({ type: 'blob' })
    .then((blob) => {
      triggerDownload(blob, 'b2c_projects.zip');
      setStatus('ZIP archive ready for download.');
    })
    .catch((error) => setStatus(`Failed to build ZIP archive: ${error.message}`, true));
}

function triggerDownload(data, filename) {
  const blob = data instanceof Blob ? data : new Blob([data]);
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle('error', Boolean(isError));
}

function cleanCell(value) {
  if (value == null) {
    return '';
  }
  if (typeof value === 'number') {
    return value.toString();
  }
  return String(value).trim();
}

function normalizeSpaces(value) {
  return value ? value.replace(/\s+/g, ' ').trim() : '';
}

function parseDuration(value) {
  const normalized = cleanCell(value).replace(/\s+/g, '').replace(',', '.');
  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : 0;
}

function createWorkbookFromRecords(header, records, { includeProject, includeSubProject, sheetName }) {
  const rows = buildSheetRows(header, records, includeProject, includeSubProject);
  const workbook = XLSX.utils.book_new();
  const safeSheet = sanitizeSheetName(sheetName);
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, worksheet, safeSheet);
  return workbook;
}

function buildSheetRows(header, records, includeProject, includeSubProject) {
  const effectiveHeader = [
    ...header,
    ...(includeProject ? ['Project'] : []),
    ...(includeSubProject ? ['Sub Project'] : [])
  ];
  const rows = [effectiveHeader];

  for (const record of records) {
    const values = record.values.slice();
    values[2] = formatDurationForExport(record.duration, record.values[2]);
    const row = [
      ...values,
      ...(includeProject ? [record.project] : []),
      ...(includeSubProject ? [record.subProject || ''] : [])
    ];
    rows.push(row);
  }

  return rows;
}

function workbookToBlob(workbook) {
  const buffer = workbookToArrayBuffer(workbook);
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

function workbookToArrayBuffer(workbook) {
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
}

function sanitizeSheetName(name) {
  const cleaned = (name || 'Sheet1').replace(/[\[\]\*\/\\\?:]/g, ' ').trim();
  const fallback = cleaned || 'Sheet1';
  return fallback.slice(0, 31);
}

function formatDurationForExport(duration, originalValue) {
  if (Number.isFinite(duration)) {
    if (typeof originalValue === 'string') {
      const match = originalValue.replace(/\s+/g, '').match(/[,\.](\d+)/);
      if (match && match[1]) {
        return duration.toFixed(match[1].length);
      }
    }
    return duration.toString();
  }
  const cleaned = cleanCell(originalValue);
  return cleaned ? cleaned.replace(',', '.') : '0';
}

function formatHours(value) {
  const numeric = Number.isFinite(value) ? value : 0;
  const rounded = Math.round(numeric * 100) / 100;
  const fraction = Math.abs(rounded % 1);
  const options = fraction === 0
    ? { minimumFractionDigits: 0, maximumFractionDigits: 0 }
    : { minimumFractionDigits: 2, maximumFractionDigits: 2 };
  return rounded.toLocaleString(undefined, options);
}

function escapeCsv(value) {
  const text = value == null ? '' : String(value);
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function toSafeFilename(name) {
  return name
    .replace(/[\\/:"*?<>|]+/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase();
}
