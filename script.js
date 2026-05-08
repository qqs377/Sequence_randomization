/* ── State ── */
let originalRows = [];      // Array of row-objects keyed by column name
let headers = [];           // Original column headers from file
let selectedCols = new Set();
let currentRows = [];       // Currently displayed (may be randomized)
let isRandomized = false;

/* ── DOM refs ── */
const fileInput       = document.getElementById('fileInput');
const dropZone        = document.getElementById('dropZone');
const fileNameEl      = document.getElementById('fileName');
const columnChips     = document.getElementById('columnChips');
const btnSelectAll    = document.getElementById('btnSelectAll');
const btnDeselectAll  = document.getElementById('btnDeselectAll');
const btnApplyColumns = document.getElementById('btnApplyColumns');
const tableHead       = document.getElementById('tableHead');
const tableBody       = document.getElementById('tableBody');
const btnRandomize    = document.getElementById('btnRandomize');
const btnReset        = document.getElementById('btnReset');
const rowCountEl      = document.getElementById('rowCount');
const btnConfirm      = document.getElementById('btnConfirm');
const confirmHint     = document.getElementById('confirmHint');
const btnExport       = document.getElementById('btnExport');
const btnStartOver    = document.getElementById('btnStartOver');

const sectionUpload   = document.getElementById('section-upload');
const sectionColumns  = document.getElementById('section-columns');
const sectionPreview  = document.getElementById('section-preview');
const sectionExport   = document.getElementById('section-export');

/* ── File upload ── */
fileInput.addEventListener('change', e => handleFile(e.target.files[0]));

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

function handleFile(file) {
  if (!file) return;
  fileNameEl.textContent = '📄 ' + file.name;

  const reader = new FileReader();
  reader.onload = evt => {
    const data = new Uint8Array(evt.target.result);
    const wb = XLSX.read(data, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (!json.length) { alert('The sheet appears to be empty.'); return; }

    headers = Object.keys(json[0]);
    originalRows = json;
    selectedCols = new Set(headers);

    buildChips();
    show(sectionColumns);
  };
  reader.readAsArrayBuffer(file);
}

/* ── Column chips ── */
function buildChips() {
  columnChips.innerHTML = '';
  headers.forEach(col => {
    const btn = document.createElement('button');
    btn.className = 'chip active';
    btn.textContent = col;
    btn.dataset.col = col;
    btn.addEventListener('click', () => {
      if (selectedCols.has(col)) {
        selectedCols.delete(col);
        btn.classList.remove('active');
      } else {
        selectedCols.add(col);
        btn.classList.add('active');
      }
    });
    columnChips.appendChild(btn);
  });
}

btnSelectAll.addEventListener('click', () => {
  headers.forEach(col => selectedCols.add(col));
  columnChips.querySelectorAll('.chip').forEach(c => c.classList.add('active'));
});

btnDeselectAll.addEventListener('click', () => {
  selectedCols.clear();
  columnChips.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
});

btnApplyColumns.addEventListener('click', () => {
  if (!selectedCols.size) { alert('Please select at least one column.'); return; }
  isRandomized = false;
  currentRows = originalRows.map((row, i) => ({ __origOrder: i + 1, ...row }));
  renderTable(currentRows, false);
  show(sectionPreview);
  btnConfirm.disabled = true;
  confirmHint.textContent = 'Randomize first to enable export.';
});

/* ── Table rendering ── */
function renderTable(rows, randomized) {
  const cols = headers.filter(h => selectedCols.has(h));

  // Header
  tableHead.innerHTML = '';
  const tr = document.createElement('tr');
  addTH(tr, 'New Order', 'col-new-order');
  addTH(tr, 'Original Order', 'col-orig-order');
  cols.forEach(c => addTH(tr, c));
  tableHead.appendChild(tr);

  // Rows
  tableBody.innerHTML = '';
  rows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    addTD(tr, idx + 1, 'col-new-order');
    addTD(tr, row.__origOrder, 'col-orig-order');
    cols.forEach(c => addTD(tr, row[c] ?? ''));
    tableBody.appendChild(tr);
  });

  rowCountEl.textContent = `${rows.length} row${rows.length !== 1 ? 's' : ''}`;
}

function addTH(tr, text, cls) {
  const th = document.createElement('th');
  th.textContent = text;
  if (cls) th.className = cls;
  tr.appendChild(th);
}

function addTD(tr, text, cls) {
  const td = document.createElement('td');
  td.textContent = text;
  if (cls) td.className = cls;
  tr.appendChild(td);
}

/* ── Randomize ── */
btnRandomize.addEventListener('click', () => {
  currentRows = shuffle([...originalRows.map((row, i) => ({ __origOrder: i + 1, ...row }))]);
  isRandomized = true;
  renderTable(currentRows, true);
  btnConfirm.disabled = false;
  confirmHint.textContent = 'Looks good? Confirm to proceed to export.';
});

btnReset.addEventListener('click', () => {
  isRandomized = false;
  currentRows = originalRows.map((row, i) => ({ __origOrder: i + 1, ...row }));
  renderTable(currentRows, false);
  btnConfirm.disabled = true;
  confirmHint.textContent = 'Randomize first to enable export.';
});

function shuffle(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

/* ── Confirm ── */
btnConfirm.addEventListener('click', () => {
  show(sectionExport);
});

/* ── Export ── */
btnExport.addEventListener('click', () => {
  const cols = headers.filter(h => selectedCols.has(h));

  // Build export data with "New Order" and "Original Order" first
  const exportData = currentRows.map((row, idx) => {
    const obj = { 'New Order': idx + 1, 'Original Order': row.__origOrder };
    cols.forEach(c => { obj[c] = row[c] ?? ''; });
    return obj;
  });

  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Randomized');

  // Style header row width hints
  const colWidths = Object.keys(exportData[0]).map(k => ({ wch: Math.max(k.length + 2, 10) }));
  ws['!cols'] = colWidths;

  XLSX.writeFile(wb, 'randomized_output.xlsx');
});

/* ── Start over ── */
btnStartOver.addEventListener('click', () => {
  originalRows = [];
  headers = [];
  selectedCols.clear();
  currentRows = [];
  isRandomized = false;
  fileInput.value = '';
  fileNameEl.textContent = '';
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';
  columnChips.innerHTML = '';
  sectionColumns.classList.add('hidden');
  sectionPreview.classList.add('hidden');
  sectionExport.classList.add('hidden');
});

/* ── Helpers ── */
function show(section) {
  [sectionColumns, sectionPreview, sectionExport].forEach(s => {
    if (s !== section) s.classList.add('hidden');
  });
  section.classList.remove('hidden');
  section.scrollIntoView({ behavior: 'smooth', block: 'start' });
}
