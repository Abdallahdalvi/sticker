const state = { rows: [], rowIndex: 0 };

const fileInput = document.getElementById('data-file');
const uploadStatus = document.getElementById('upload-status');
const errorPanel = document.getElementById('error-panel');
const errorSummary = document.getElementById('error-summary');
const errorRows = document.getElementById('error-rows');
const exportCard = document.getElementById('export-card');
const startRecord = document.getElementById('start-record');
const endRecord = document.getElementById('end-record');
const printerDpi = document.getElementById('printer-dpi');
const pageFormat = document.getElementById('page-format');
const generateButton = document.getElementById('generate-pdf');
const printWarning = document.getElementById('print-warning');
const previewImage = document.getElementById('preview-image');
const previewTitle = document.getElementById('preview-title');
const rowNav = document.getElementById('row-nav');
const rowLabel = document.getElementById('row-label');
const summaryValues = document.querySelectorAll('#data-summary strong');

function setBusy(busy, text = '') {
  fileInput.disabled = busy;
  generateButton.disabled = busy || !state.rows.length;
  if (text) {
    uploadStatus.textContent = text;
    uploadStatus.className = 'status info';
  }
}

function clearErrors() {
  errorPanel.classList.add('hidden');
  errorRows.innerHTML = '';
  errorSummary.textContent = '';
}

function showErrors(payload) {
  const detail = payload?.detail || payload || {};
  const errors = detail.errors || [{ row: 0, field: 'File', message: detail.message || 'Unexpected error.' }];
  errorSummary.textContent = detail.message || `${errors.length} validation issue(s) were found.`;
  errorRows.innerHTML = '';
  errors.slice(0, 200).forEach(issue => {
    const row = document.createElement('tr');
    [issue.row || '—', issue.field || '—', issue.message || 'Unknown issue'].forEach(value => {
      const cell = document.createElement('td');
      cell.textContent = value;
      row.appendChild(cell);
    });
    errorRows.appendChild(row);
  });
  errorPanel.classList.remove('hidden');
  errorPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function parseError(response) {
  try { return await response.json(); }
  catch { return { detail: { message: await response.text(), errors: [] } }; }
}

function updateRangeSummary() {
  if (!state.rows.length) return;
  const start = Math.max(1, Number(startRecord.value) || 1);
  const end = Math.min(state.rows.length, Number(endRecord.value) || state.rows.length);
  const records = end >= start ? end - start + 1 : 0;
  const pages = pageFormat.value === 'a4' ? Math.ceil(records / 21) : records;
  summaryValues[0].textContent = records.toLocaleString();
  summaryValues[1].textContent = pages.toLocaleString();
  summaryValues[2].textContent = '3';
  printWarning.textContent = pageFormat.value === 'a4'
    ? 'A4: 21 stickers per sheet (7 × 3) with 4 mm cutting gaps. Print at 100% / Actual Size.'
    : 'Print at 100% / Actual Size. Do not use “Fit to page”.';
}

async function refreshPreview() {
  if (!state.rows.length) return;
  rowLabel.textContent = `Row ${state.rowIndex + 1} / ${state.rows.length}`;
  previewTitle.textContent = state.rows[state.rowIndex]['Device ID'];
  const response = await fetch('/api/preview', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ row: state.rows[state.rowIndex], printerDpi: Number(printerDpi.value) }),
  });
  if (!response.ok) {
    showErrors(await parseError(response));
    return;
  }
  const data = await response.json();
  previewImage.src = data.svg;
}

fileInput.addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  clearErrors();
  setBusy(true, `Checking ${file.name}…`);
  const body = new FormData();
  body.append('file', file);
  try {
    const response = await fetch('/api/upload', { method: 'POST', body });
    if (!response.ok) {
      state.rows = [];
      showErrors(await parseError(response));
      uploadStatus.textContent = 'Upload rejected. Correct the listed rows and try again.';
      uploadStatus.className = 'status error';
      return;
    }
    const data = await response.json();
    state.rows = data.rows;
    state.rowIndex = 0;
    startRecord.value = 1;
    endRecord.value = data.count;
    startRecord.max = data.count;
    endRecord.max = data.count;
    [startRecord, endRecord, printerDpi, pageFormat, generateButton].forEach(control => { control.disabled = false; });
    exportCard.classList.remove('muted');
    rowNav.classList.remove('hidden');
    uploadStatus.textContent = `${data.count.toLocaleString()} records validated. Identifiers remain exact text.`;
    uploadStatus.className = 'status success';
    updateRangeSummary();
    await refreshPreview();
  } catch (error) {
    showErrors({ detail: { message: 'Could not reach the local generator.', errors: [{ row: 0, field: 'App', message: error.message }] } });
  } finally {
    setBusy(false);
    fileInput.value = '';
  }
});

document.getElementById('prev-row').addEventListener('click', async () => {
  if (!state.rows.length) return;
  state.rowIndex = Math.max(0, state.rowIndex - 1);
  await refreshPreview();
});

document.getElementById('next-row').addEventListener('click', async () => {
  if (!state.rows.length) return;
  state.rowIndex = Math.min(state.rows.length - 1, state.rowIndex + 1);
  await refreshPreview();
});

[startRecord, endRecord].forEach(control => control.addEventListener('input', updateRangeSummary));
printerDpi.addEventListener('change', refreshPreview);
pageFormat.addEventListener('change', updateRangeSummary);

generateButton.addEventListener('click', async () => {
  if (!state.rows.length) return;
  clearErrors();
  const start = Number(startRecord.value);
  const end = Number(endRecord.value);
  generateButton.disabled = true;
  generateButton.textContent = 'Generating…';
  try {
    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        rows: state.rows,
        start,
        end,
        printerDpi: Number(printerDpi.value),
        pageFormat: pageFormat.value,
      }),
    });
    if (!response.ok) {
      showErrors(await parseError(response));
      return;
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    const suffix = pageFormat.value === 'a4' ? '_a4' : '';
    anchor.download = `stickers_${start}-${end}${suffix}.pdf`;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    const mode = pageFormat.value === 'a4' ? 'A4 sticker sheet' : 'sticker';
    uploadStatus.textContent = `${(end - start + 1).toLocaleString()} ${mode} PDF generated successfully.`;
    uploadStatus.className = 'status success';
  } catch (error) {
    showErrors({ detail: { message: 'PDF generation failed.', errors: [{ row: 0, field: 'PDF', message: error.message }] } });
  } finally {
    generateButton.disabled = false;
    generateButton.textContent = 'Generate sticker PDF';
  }
});
