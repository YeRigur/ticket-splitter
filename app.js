import { filterDefinitions, compileFilters, realProjects } from './filters.js';

const compiledFilters = compileFilters(filterDefinitions);
const projectOrder = [...filterDefinitions.map((entry) => entry.name), 'Other'];

const fileInput = document.getElementById('csv-input');
const hasHeaderCheckbox = document.getElementById('has-header');
const statusEl = document.getElementById('status');
const resultsSection = document.getElementById('results');
const summaryBody = document.getElementById('summary-body');
const zipButton = document.getElementById('download-zip');
const combinedButton = document.getElementById('download-combined');
const overallHoursEl = document.getElementById('overall-hours');

let state = null;

fileInput.addEventListener('change', handleFileSelection);
zipButton.addEventListener('click', downloadZipArchive);
combinedButton.addEventListener('click', downloadCombinedFile);
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

  if (!file.name.toLowerCase().endsWith('.csv')) {
    setStatus('Please select a CSV file.', true);
    return;
  }

  setStatus('Reading file...');
  Papa.parse(file, {
    skipEmptyLines: 'greedy',
    complete: (result) => processData(result.data),
    error: (error) => setStatus(`Failed to read file: ${error.message}`, true)
  });
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

  const grouped = new Map();
  for (const row of dataRows) {
    const comment = normalizeSpaces(row[1]);
    const duration = parseDuration(row[2]);
    const project = classify(comment);
    const record = { project, values: row, duration };
    if (!grouped.has(project)) {
      grouped.set(project, []);
    }
    grouped.get(project).push(record);
  }

  state = { header, grouped, total: dataRows.length };
  resultsSection.classList.remove('hidden');
  setStatus(`Processed rows: ${dataRows.length}`);
  updateSummary();
}

function classify(comment) {
  for (const { name, regexes } of compiledFilters) {
    if (regexes.some((regex) => regex.test(comment))) {
      return name;
    }
  }
  return 'Other';
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
  zipButton.disabled = !hasAnyData || typeof JSZip === 'undefined';
  combinedButton.disabled = !realProjects.some((project) => (state.grouped.get(project) ?? []).length > 0);
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
  button.textContent = 'CSV';
  button.dataset.project = project;
  actionsCell.appendChild(button);
  row.appendChild(actionsCell);

  return row;
}

function downloadProject(project) {
  if (!state) {
    return;
  }
  const entries = state.grouped.get(project);
  if (!entries?.length) {
    setStatus(`No rows found for project "${project}".`, true);
    return;
  }
  const csv = createCsv(state.header, entries);
  triggerDownload(csv, `${toSafeFilename(project) || 'project'}.csv`);
  setStatus(`CSV for project "${project}" ready for download.`);
}

function downloadCombinedFile() {
  if (!state) {
    return;
  }
  const combined = [];
  for (const project of realProjects) {
    const entries = state.grouped.get(project);
    if (entries?.length) {
      combined.push(...entries);
    }
  }
  if (!combined.length) {
    setStatus('No rows available for the combined report.', true);
    return;
  }
  const csv = createCsv(state.header, combined, { includeProject: true });
  triggerDownload(csv, 'all_b2c_projects.csv');
  setStatus('Combined CSV ready for download.');
}

function downloadZipArchive() {
  if (!state) {
    return;
  }
  if (typeof JSZip === 'undefined') {
    setStatus('JSZip failed to load. Check the CDN script tag.', true);
    return;
  }

  const zip = new JSZip();
  let filesAdded = 0;

  for (const [project, entries] of state.grouped.entries()) {
    if (!entries.length) {
      continue;
    }
    const csv = createCsv(state.header, entries);
    zip.file(`${toSafeFilename(project) || 'project'}.csv`, csv);
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

function createCsv(header, records, { includeProject = false } = {}) {
  const effectiveHeader = includeProject ? ['Project', ...header] : header;
  const lines = [effectiveHeader.map(escapeCsv).join(',')];

  for (const record of records) {
    const row = includeProject ? [record.project, ...record.values] : record.values;
    lines.push(row.map(escapeCsv).join(','));
  }

  return lines.join('\n');
}

function triggerDownload(data, filename) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: 'text/csv;charset=utf-8;' });
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
