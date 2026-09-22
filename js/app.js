// ==========================================================
// AETHERIA AutoSend Pro — WhatsApp Automation Engine
// Persistent IndexedDB Storage · History · Template Library
// ==========================================================

// ========== APPLICATION STATE ==========
let currentWorkbook = null;
let currentSheetName = '';
let currentFileName = '';
let activeCampaign = null;     // Active campaign object saved in DB

let parsedData = [];          // Raw rows
let normalizedData = [];      // Rows with validated phone numbers
let columns = [];             // Header columns
let phoneColumn = '';         // Selected phone column
let messageTemplate = '';     // Template with {{placeholders}}
let activeMode = 'quick';     // 'quick' or 'auto'
let activeView = 'workspace'; // 'workspace', 'history', 'templates'

// Queue state
let currentQueueIndex = 0;
let autoTimer = null;
let isAutoRunning = false;
let isAutoPaused = false;
const dispatchLogs = [];

// Pagination
let currentPage = 1;
const rowsPerPage = 25;

// ========== INITIALIZATION ==========
document.addEventListener('DOMContentLoaded', async () => {
  initUploadZone();
  initTemplateEditor();
  initControls();
  initPagination();

  // Initialize DB and load badges/dropdowns
  await window.appDB.readyPromise;
  await refreshDatabaseBadges();
  await populateTemplateDropdown();
});

// ========== VIEW NAVIGATION ==========
function switchView(view) {
  activeView = view;

  document.getElementById('navTabWorkspace').classList.toggle('active', view === 'workspace');
  document.getElementById('navTabHistory').classList.toggle('active', view === 'history');
  document.getElementById('navTabTemplates').classList.toggle('active', view === 'templates');

  document.getElementById('workspaceView').style.display = view === 'workspace' ? 'block' : 'none';
  document.getElementById('historyView').style.display = view === 'history' ? 'block' : 'none';
  document.getElementById('templatesView').style.display = view === 'templates' ? 'block' : 'none';

  if (view === 'history') {
    renderHistoryTable();
  } else if (view === 'templates') {
    renderTemplatesGrid();
  }
}

async function refreshDatabaseBadges() {
  const camps = await window.appDB.getAllCampaigns();
  const tpls = await window.appDB.getAllTemplates();

  document.getElementById('campCountBadge').textContent = camps.length;
  document.getElementById('tplCountBadge').textContent = tpls.length;
}

// ========== STEP 1: FILE UPLOADER & MULTI-FORMAT PARSER ==========
function initUploadZone() {
  const zone = document.getElementById('uploadZone');
  const input = document.getElementById('fileInput');
  const removeBtn = document.getElementById('removeFile');
  const sheetSelect = document.getElementById('sheetSelect');

  zone.addEventListener('click', () => input.click());

  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });

  zone.addEventListener('dragleave', () => {
    zone.classList.remove('drag-over');
  });

  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      handleIncomingFile(e.dataTransfer.files[0]);
    }
  });

  input.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleIncomingFile(e.target.files[0]);
    }
  });

  removeBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    resetApplication();
  });

  sheetSelect.addEventListener('change', (e) => {
    if (currentWorkbook) {
      loadSheetData(e.target.value);
    }
  });
}

function handleIncomingFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const validExtensions = ['xlsx', 'xls', 'csv', 'pdf', 'docx'];

  if (!validExtensions.includes(ext)) {
    showToast('Unsupported file type. Please upload .xlsx, .csv, .pdf, or .docx', 'error');
    return;
  }

  currentFileName = file.name;
  document.getElementById('fileInfo').style.display = 'block';
  document.getElementById('fileName').textContent = file.name;
  document.getElementById('fileSize').textContent = `(${formatFileSize(file.size)})`;

  document.getElementById('loadingBar').style.display = 'flex';
  document.getElementById('uploadZone').style.display = 'none';

  setTimeout(() => {
    switch (ext) {
      case 'xlsx':
      case 'xls':
        parseExcelFile(file);
        break;
      case 'csv':
        parseCSVFile(file);
        break;
      case 'pdf':
        parsePDFFile(file);
        break;
      case 'docx':
        parseDOCXFile(file);
        break;
    }
  }, 100);
}

// --- Excel Parser ---
function parseExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      currentWorkbook = XLSX.read(data, { type: 'array', cellDates: true });

      const sheetNames = currentWorkbook.SheetNames;
      const sheetSelect = document.getElementById('sheetSelect');
      const sheetWrapper = document.getElementById('sheetSelectorWrapper');

      if (sheetNames.length > 1) {
        sheetSelect.innerHTML = '';
        sheetNames.forEach((name, idx) => {
          const opt = document.createElement('option');
          opt.value = name;
          opt.textContent = `${name} (${idx + 1}/${sheetNames.length})`;
          sheetSelect.appendChild(opt);
        });
        sheetWrapper.style.display = 'block';
      } else {
        sheetWrapper.style.display = 'none';
      }

      loadSheetData(sheetNames[0]);
    } catch (err) {
      showToast('Error parsing Excel: ' + err.message, 'error');
      hideLoading();
    }
  };
  reader.readAsArrayBuffer(file);
}

function loadSheetData(sheetName) {
  currentSheetName = sheetName;
  const sheet = currentWorkbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });

  if (!rawRows || rawRows.length === 0) {
    showToast(`No records found in sheet "${sheetName}".`, 'error');
    hideLoading();
    return;
  }

  const cleanedRows = rawRows.map(row => {
    const cleaned = {};
    for (const k of Object.keys(row)) {
      cleaned[k.trim()] = cleanFieldValue(row[k]);
    }
    return cleaned;
  });

  onDataLoaded(cleanedRows);
}

// --- CSV Parser ---
function parseCSVFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const text = e.target.result;
      const workbook = XLSX.read(text, { type: 'string', raw: false });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });

      if (!rawRows || rawRows.length === 0) {
        showToast('No records found in CSV file.', 'error');
        hideLoading();
        return;
      }

      const cleanedRows = rawRows.map(row => {
        const cleaned = {};
        for (const k of Object.keys(row)) {
          cleaned[k.trim()] = cleanFieldValue(row[k]);
        }
        return cleaned;
      });

      onDataLoaded(cleanedRows);
    } catch (err) {
      showToast('Error reading CSV: ' + err.message, 'error');
      hideLoading();
    }
  };
  reader.readAsText(file);
}

// --- PDF Parser ---
function parsePDFFile(file) {
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const typedArray = new Uint8Array(e.target.result);
      pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

      const pdf = await pdfjsLib.getDocument(typedArray).promise;
      let fullText = '';

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        fullText += content.items.map(item => item.str).join(' ') + '\n';
      }

      const extracted = extractRecordsFromText(fullText);
      if (extracted.length === 0) {
        showToast('Could not find structured contacts in PDF. Please try Excel or CSV format.', 'error');
        hideLoading();
        return;
      }

      onDataLoaded(extracted);
    } catch (err) {
      showToast('Error reading PDF: ' + err.message, 'error');
      hideLoading();
    }
  };
  reader.readAsArrayBuffer(file);
}

// --- DOCX Parser ---
function parseDOCXFile(file) {
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const arrayBuffer = e.target.result;
      const result = await mammoth.extractRawText({ arrayBuffer });
      const extracted = extractRecordsFromText(result.value);

      if (extracted.length === 0) {
        showToast('Could not find structured contacts in Word document. Please try Excel or CSV format.', 'error');
        hideLoading();
        return;
      }

      onDataLoaded(extracted);
    } catch (err) {
      showToast('Error reading DOCX: ' + err.message, 'error');
      hideLoading();
    }
  };
  reader.readAsArrayBuffer(file);
}

function extractRecordsFromText(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length < 2) return [];

  for (const sep of ['\t', '|', ';', ',']) {
    const first = lines[0];
    if (first.includes(sep)) {
      const headers = first.split(sep).map(h => h.trim()).filter(h => h);
      if (headers.length >= 2) {
        const rows = [];
        for (let i = 1; i < lines.length; i++) {
          const cells = lines[i].split(sep).map(c => c.trim());
          if (cells.length >= headers.length - 1) {
            const r = {};
            headers.forEach((h, idx) => {
              r[h] = cleanFieldValue(cells[idx] || '');
            });
            rows.push(r);
          }
        }
        if (rows.length > 0) return rows;
      }
    }
  }

  const phonePattern = /(\+?\d{1,3}[-.\s]?)?[6-9]\d{9}/g;
  const records = [];
  lines.forEach(line => {
    const match = line.match(phonePattern);
    if (match) {
      const remaining = line.replace(phonePattern, '').trim();
      const parts = remaining.split(/[,-|]/).map(p => p.trim()).filter(p => p);
      records.push({
        'Name': parts[0] || 'Customer',
        'Phone': match[0],
        'Details': parts.slice(1).join(', ') || ''
      });
    }
  });

  return records;
}

// ========== STEP 2: PHONE NUMBER SANITIZATION & RECOGNITION ==========

function cleanFieldValue(val) {
  if (val === null || val === undefined) return '';
  let s = String(val).trim();

  // Decode scientific notation e.g. 9.87654E+09
  if (/^[0-9]+(\.[0-9]+)?[eE]\+[0-9]+$/.test(s)) {
    try {
      s = BigInt(Math.round(Number(s))).toString();
    } catch (e) {
      s = Number(s).toFixed(0);
    }
  }

  s = s.replace(/\.0+$/, '');
  return s;
}

function sanitizePhone(raw, defaultCountryCode = '91') {
  if (raw === null || raw === undefined) {
    return { valid: false, clean: '', original: '', reason: 'Empty value' };
  }

  let s = cleanFieldValue(raw);
  const original = s;

  if (!s) {
    return { valid: false, clean: '', original: '', reason: 'Empty' };
  }

  const hasPlus = s.startsWith('+');
  let digits = s.replace(/\D/g, '');
  digits = digits.replace(/^0+/, '');

  let cc = String(defaultCountryCode).replace(/\D/g, '') || '91';
  let clean = '';

  if (hasPlus) {
    clean = digits;
  } else if (digits.length === 10) {
    clean = cc + digits;
  } else if (digits.length === 12 && digits.startsWith(cc)) {
    clean = digits;
  } else if (digits.length === 11 && digits.startsWith(cc.slice(0, 1))) {
    clean = cc + digits.slice(1);
  } else if (digits.length > 10 && digits.startsWith(cc)) {
    clean = digits;
  } else if (digits.length >= 7 && digits.length <= 15) {
    clean = cc + digits;
  } else {
    clean = digits;
  }

  const isValid = clean.length >= 10 && clean.length <= 15;
  const reason = isValid ? 'Valid' : `Invalid length (${clean.length} digits)`;

  return { valid: isValid, clean, original, reason };
}

function autoDetectPhoneColumn(cols, rows) {
  const phoneKeywords = ['phone', 'mobile', 'contact', 'whatsapp', 'cell', 'number', 'tel', 'mob', 'ph'];

  for (const col of cols) {
    const lower = col.toLowerCase();
    if (phoneKeywords.some(k => lower.includes(k))) {
      return col;
    }
  }

  const scanLimit = Math.min(15, rows.length);
  let bestCol = cols[0];
  let maxScore = -1;

  for (const col of cols) {
    let score = 0;
    for (let i = 0; i < scanLimit; i++) {
      const val = cleanFieldValue(rows[i][col]).replace(/\D/g, '');
      if (val.length >= 10 && val.length <= 13) {
        score++;
      }
    }
    if (score > maxScore) {
      maxScore = score;
      bestCol = col;
    }
  }

  return bestCol;
}

// ========== DATA LOAD & DATABASE AUTO-SAVE ==========
async function onDataLoaded(rows) {
  parsedData = rows;
  columns = Object.keys(rows[0]);

  hideLoading();

  phoneColumn = autoDetectPhoneColumn(columns, parsedData);
  recomputeNormalizedData();

  setupPhoneDropdown();
  setupRangeInputs();
  renderStatsBanner();
  renderDataTable();
  renderVariableTags();
  renderLivePreviews();
  updateQuickQueueCard();

  document.getElementById('step2').style.display = 'block';
  document.getElementById('step3').style.display = 'block';
  document.getElementById('step4').style.display = 'block';

  // Automatically save new campaign into IndexedDB
  await saveCurrentCampaign(false);
  await refreshDatabaseBadges();

  showToast(`✅ Loaded ${parsedData.length} records & saved to database!`, 'success');

  setTimeout(() => {
    document.getElementById('step2').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 200);
}

function recomputeNormalizedData() {
  const cc = document.getElementById('countryCode').value;

  normalizedData = parsedData.map((row, idx) => {
    const rawVal = row[phoneColumn];
    const phoneInfo = sanitizePhone(rawVal, cc);
    return {
      _index: idx,
      _phoneInfo: phoneInfo,
      _status: row._status || (phoneInfo.valid ? 'ready' : 'invalid'),
      ...row
    };
  });
}

function renderStatsBanner() {
  const total = normalizedData.length;
  const valid = normalizedData.filter(r => r._phoneInfo.valid).length;
  const invalid = total - valid;

  document.getElementById('statTotal').textContent = total;
  document.getElementById('statValid').textContent = valid;
  document.getElementById('statInvalid').textContent = invalid;
  document.getElementById('rowCount').textContent = `${total} Contacts Loaded`;
}

function setupPhoneDropdown() {
  const select = document.getElementById('phoneColumnSelect');
  select.innerHTML = '';

  columns.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col;
    opt.textContent = col;
    if (col === phoneColumn) opt.selected = true;
    select.appendChild(opt);
  });

  select.onchange = (e) => {
    phoneColumn = e.target.value;
    recomputeNormalizedData();
    renderStatsBanner();
    renderDataTable();
    renderLivePreviews();
    updateQuickQueueCard();
    saveCurrentCampaign(false);
    showToast(`Updated phone column to "${phoneColumn}"`, 'info');
  };
}

function setupRangeInputs() {
  const fromInput = document.getElementById('rangeFrom');
  const toInput = document.getElementById('rangeTo');
  const startInput = document.getElementById('startSrNo');
  const endInput = document.getElementById('endSrNo');
  const bdInfo = document.getElementById('bdDatasetInfo');

  const total = normalizedData.length || parsedData.length || 1;

  if (fromInput) { fromInput.value = 1; fromInput.max = total; fromInput.onchange = () => syncStartingPosition(fromInput.value); }
  if (toInput) { toInput.value = total; toInput.max = total; toInput.onchange = () => syncEndingPosition(toInput.value); }
  if (startInput) { startInput.value = 1; startInput.max = total; }
  if (endInput) { endInput.value = Math.min(total, 50); endInput.max = total; }
  if (bdInfo) { bdInfo.textContent = `Total dataset: ${total} contacts`; }

  updateRangeStatus();
}

function syncStartingPosition(val) {
  let start = parseInt(val) || 1;
  const total = normalizedData.length;
  if (start < 1) start = 1;
  if (start > total) start = total;

  document.getElementById('startSrNo').value = start;
  document.getElementById('rangeFrom').value = start;

  let end = parseInt(document.getElementById('endSrNo').value) || total;
  if (end < start) {
    end = Math.min(total, start + 49);
    document.getElementById('endSrNo').value = end;
    document.getElementById('rangeTo').value = end;
  }

  currentQueueIndex = 0;
  updateQuickQueueCard();
  updateRangeStatus();
  showToast(`🎯 Queue will start from Sr. No. ${start}`, 'info');
}

function syncEndingPosition(val) {
  let end = parseInt(val) || normalizedData.length;
  const total = normalizedData.length;
  let start = parseInt(document.getElementById('startSrNo').value) || 1;

  if (end < start) end = start;
  if (end > total) end = total;

  document.getElementById('endSrNo').value = end;
  document.getElementById('rangeTo').value = end;

  currentQueueIndex = 0;
  updateQuickQueueCard();
  updateRangeStatus();
}

function setBatchSize(size) {
  const total = normalizedData.length;
  if (total === 0) return;

  let start = parseInt(document.getElementById('startSrNo').value) || 1;
  let end = Math.min(total, start + size - 1);

  document.getElementById('endSrNo').value = end;
  document.getElementById('rangeTo').value = end;

  currentQueueIndex = 0;
  updateQuickQueueCard();
  updateRangeStatus();
  showToast(`📦 Set batch: Sr. No. ${start} to ${end} (${end - start + 1} contacts)`, 'success');
}

function resetRangeToAll() {
  const total = normalizedData.length;
  if (total === 0) return;

  document.getElementById('startSrNo').value = 1;
  document.getElementById('rangeFrom').value = 1;
  document.getElementById('endSrNo').value = total;
  document.getElementById('rangeTo').value = total;

  currentQueueIndex = 0;
  updateQuickQueueCard();
  updateRangeStatus();
  showToast(`Targeting all ${total} contacts`, 'info');
}

function jumpToFirstUnsent() {
  const total = normalizedData.length;
  if (total === 0) {
    showToast('No data loaded yet.', 'error');
    return;
  }

  // Find first row that is not sent
  const unsentIndex = normalizedData.findIndex(r => r._status !== 'sent');

  if (unsentIndex === -1) {
    showToast('🎉 All contacts in this dataset are already marked as Sent!', 'success');
    return;
  }

  const startSr = unsentIndex + 1;
  const endSr = Math.min(total, startSr + 49);

  document.getElementById('startSrNo').value = startSr;
  document.getElementById('rangeFrom').value = startSr;
  document.getElementById('endSrNo').value = endSr;
  document.getElementById('rangeTo').value = endSr;

  currentQueueIndex = 0;
  updateQuickQueueCard();
  updateRangeStatus();

  const contactName = normalizedData[unsentIndex][columns[0]] || `Contact ${startSr}`;
  showToast(`⏩ Jumped to first unsent: Sr. No. ${startSr} (${contactName})`, 'success');
}

function startAutomationFromRow(rowIndex) {
  const total = normalizedData.length;
  const startSr = rowIndex + 1;
  const endSr = Math.min(total, startSr + 49);

  document.getElementById('startSrNo').value = startSr;
  document.getElementById('rangeFrom').value = startSr;
  document.getElementById('endSrNo').value = endSr;
  document.getElementById('rangeTo').value = endSr;

  currentQueueIndex = 0;
  updateQuickQueueCard();
  updateRangeStatus();

  const step4 = document.getElementById('step4');
  if (step4) {
    step4.style.display = 'block';
    step4.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  showToast(`🎯 Automation will start from Sr. No. ${startSr}!`, 'success');
}

function updateRangeStatus() {
  const start = parseInt(document.getElementById('startSrNo').value) || 1;
  const end = parseInt(document.getElementById('endSrNo').value) || normalizedData.length || 1;
  const total = normalizedData.length || 0;
  const batchCount = Math.max(0, end - start + 1);

  const statusBar = document.getElementById('bdStatusBar');
  if (statusBar) {
    statusBar.innerHTML = `
      Targeting Sr. No. <strong>#${start}</strong> to <strong>#${end}</strong> 
      (${batchCount} contacts in this batch · Total dataset: ${total}).
    `;
  }

  const bdInfo = document.getElementById('bdDatasetInfo');
  if (bdInfo) {
    bdInfo.textContent = `Total dataset: ${total} contacts`;
  }
}

function initControls() {
  const ccInput = document.getElementById('countryCode');
  ccInput.addEventListener('input', () => {
    recomputeNormalizedData();
    renderStatsBanner();
    renderDataTable();
    renderLivePreviews();
    updateQuickQueueCard();
    saveCurrentCampaign(false);
  });

  const slider = document.getElementById('delaySlider');
  const display = document.getElementById('delayValue');
  slider.addEventListener('input', () => {
    display.textContent = slider.value + 's';
  });

  window.addEventListener('keydown', (e) => {
    if (activeView === 'workspace' && activeMode === 'quick' && (e.code === 'Space' || e.code === 'Enter')) {
      const activeTag = document.activeElement ? document.activeElement.tagName : '';
      if (activeTag !== 'TEXTAREA' && activeTag !== 'INPUT') {
        const step4 = document.getElementById('step4');
        if (step4.style.display !== 'none') {
          e.preventDefault();
          quickSendCurrent();
        }
      }
    }
  });
}

// ========== RENDER DATA TABLE ==========
function renderDataTable() {
  const thead = document.getElementById('dataTableHead');
  thead.innerHTML = `
    <tr>
      <th style="width: 50px;">#</th>
      <th style="width: 120px;">Action</th>
      <th>📱 WhatsApp No.</th>
      <th>Status</th>
      ${columns.map(c => `<th>${escapeHtml(c)}</th>`).join('')}
    </tr>
  `;

  renderTablePage();
}

function renderTablePage() {
  const tbody = document.getElementById('dataTableBody');
  const start = (currentPage - 1) * rowsPerPage;
  const end = Math.min(start + rowsPerPage, normalizedData.length);
  const totalPages = Math.ceil(normalizedData.length / rowsPerPage) || 1;

  tbody.innerHTML = '';

  for (let i = start; i < end; i++) {
    const row = normalizedData[i];
    const phone = row._phoneInfo;
    const tr = document.createElement('tr');

    let statusBadge = '';
    if (row._status === 'sent') {
      statusBadge = `<span class="badge-valid" style="background: rgba(37,211,102,0.2);">✅ Sent</span>`;
    } else if (row._status === 'skipped') {
      statusBadge = `<span class="badge-invalid" style="background: rgba(255,183,77,0.2); color: var(--warning);">⏭️ Skipped</span>`;
    } else if (phone.valid) {
      statusBadge = `<span class="badge-valid">Ready</span>`;
    } else {
      statusBadge = `<span class="badge-invalid" title="${escapeHtml(phone.reason)}">⚠️ ${escapeHtml(phone.reason)}</span>`;
    }

    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>
        <div style="display: flex; gap: 6px; align-items: center;">
          <button class="btn-test-send" onclick="sendIndividualTest(${i})" title="Test send directly to this contact">
            💬 Send
          </button>
          <button class="btn-start-here" onclick="startAutomationFromRow(${i})" title="Start automation queue from this Sr No. (#${i + 1})">
            🎯 Start
          </button>
        </div>
      </td>
      <td><strong style="color: var(--whatsapp);">+${phone.clean || '—'}</strong></td>
      <td>${statusBadge}</td>
      ${columns.map(c => `<td title="${escapeHtml(String(row[c]))}">${escapeHtml(String(row[c]))}</td>`).join('')}
    `;

    tbody.appendChild(tr);
  }

  document.getElementById('pageInfo').textContent = `Page ${currentPage} of ${totalPages}`;
  document.getElementById('prevPage').disabled = currentPage <= 1;
  document.getElementById('nextPage').disabled = currentPage >= totalPages;
}

function initPagination() {
  document.getElementById('prevPage').addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      renderTablePage();
    }
  });

  document.getElementById('nextPage').addEventListener('click', () => {
    const totalPages = Math.ceil(normalizedData.length / rowsPerPage);
    if (currentPage < totalPages) {
      currentPage++;
      renderTablePage();
    }
  });
}

// ========== STEP 3: TEMPLATE ENGINE ==========
function initTemplateEditor() {
  const textarea = document.getElementById('messageTemplate');
  textarea.addEventListener('input', (e) => {
    messageTemplate = e.target.value;
    renderLivePreviews();
    updateQuickQueueCard();
    if (activeCampaign) {
      activeCampaign.template = messageTemplate;
      window.appDB.saveCampaign(activeCampaign);
    }
  });
}

function renderVariableTags() {
  const container = document.getElementById('variableTags');
  container.innerHTML = '';

  columns.forEach(col => {
    const tag = document.createElement('span');
    tag.className = 'var-tag';
    tag.textContent = `{{${col}}}`;
    tag.addEventListener('click', () => {
      const textarea = document.getElementById('messageTemplate');
      const cursorPos = textarea.selectionStart;
      const text = textarea.value;
      const insertion = `{{${col}}}`;
      textarea.value = text.slice(0, cursorPos) + insertion + text.slice(cursorPos);
      textarea.focus();
      textarea.selectionStart = textarea.selectionEnd = cursorPos + insertion.length;

      messageTemplate = textarea.value;
      renderLivePreviews();
      updateQuickQueueCard();
    });
    container.appendChild(tag);
  });
}

function interpolate(template, row) {
  if (!template) return '';
  return template.replace(/\{\{([^}]+)\}\}/g, (match, rawKey) => {
    const key = rawKey.trim();
    const colMatch = columns.find(c => c.toLowerCase() === key.toLowerCase());
    if (colMatch && row[colMatch] !== undefined) {
      return String(row[colMatch]);
    }
    return match;
  });
}

function renderLivePreviews() {
  const container = document.getElementById('previewCards');

  if (!messageTemplate || normalizedData.length === 0) {
    container.innerHTML = '<p class="placeholder-text">Type a template to see how real contacts will see it...</p>';
    return;
  }

  const previewCount = Math.min(3, normalizedData.length);
  let html = '';

  for (let i = 0; i < previewCount; i++) {
    const row = normalizedData[i];
    const phone = row._phoneInfo;
    const msg = interpolate(messageTemplate, row);

    html += `
      <div class="preview-card">
        <div class="preview-phone">📱 +${phone.clean} (${phone.valid ? 'Ready' : '⚠️ Invalid'})</div>
        <div class="preview-msg">${escapeHtml(msg)}</div>
      </div>
    `;
  }

  container.innerHTML = html;
}

// ========== STEP 4: DISPATCH ENGINE ==========
function switchDispatchMode(mode) {
  activeMode = mode;

  const quickBtn = document.getElementById('modeQuickBtn');
  const autoBtn = document.getElementById('modeAutoBtn');
  const quickContainer = document.getElementById('modeQuickContainer');
  const autoContainer = document.getElementById('modeAutoContainer');

  if (mode === 'quick') {
    quickBtn.classList.add('active');
    autoBtn.classList.remove('active');
    quickContainer.style.display = 'block';
    autoContainer.style.display = 'none';
    updateQuickQueueCard();
  } else {
    quickBtn.classList.remove('active');
    autoBtn.classList.add('active');
    quickContainer.style.display = 'none';
    autoContainer.style.display = 'block';
  }
}

function sendIndividualTest(rowIndex) {
  const row = normalizedData[rowIndex];
  if (!row) return;

  const phone = row._phoneInfo;
  if (!phone.valid) {
    showToast(`Cannot send: ${phone.reason}`, 'error');
    return;
  }

  const template = messageTemplate.trim() || `Hello {{${columns[0]}}}, this is a test notification.`;
  const msg = interpolate(template, row);

  const url = `https://wa.me/${phone.clean}?text=${encodeURIComponent(msg)}`;
  window.open(url, 'AetheriaWhatsAppTab');

  const logEntry = {
    phone: phone.clean,
    name: row[columns[0]] || `Row ${rowIndex + 1}`,
    message: msg
  };
  addLogEntry(logEntry, 'sent', 'Single test message opened');

  row._status = 'sent';
  renderDataTable();
  if (activeCampaign) {
    window.appDB.updateRowStatus(activeCampaign.id, rowIndex, 'sent', logEntry);
  }

  showToast(`🚀 Opened WhatsApp for +${phone.clean}`, 'success');
}

function getActiveRangeData() {
  const startEl = document.getElementById('startSrNo');
  const fromEl = document.getElementById('rangeFrom');
  const endEl = document.getElementById('endSrNo');
  const toEl = document.getElementById('rangeTo');

  const fromVal = Math.max(1, parseInt(startEl?.value || fromEl?.value) || 1) - 1;
  const toVal = Math.min(normalizedData.length, parseInt(endEl?.value || toEl?.value) || normalizedData.length);
  return normalizedData.slice(fromVal, toVal);
}

function updateQuickQueueCard() {
  const range = getActiveRangeData();
  const counter = document.getElementById('quickCounter');
  const nameEl = document.getElementById('quickContactName');
  const phoneEl = document.getElementById('quickContactPhone');
  const msgEl = document.getElementById('quickContactMsg');
  const sendBtn = document.getElementById('quickSendBtn');

  if (range.length === 0 || normalizedData.length === 0) {
    nameEl.textContent = 'No records in selected range';
    phoneEl.textContent = '—';
    msgEl.textContent = 'Please adjust range above.';
    sendBtn.disabled = true;
    return;
  }

  if (currentQueueIndex >= range.length) {
    nameEl.textContent = '🎉 Batch Complete!';
    phoneEl.textContent = 'All contacts in selected range processed';
    msgEl.textContent = 'Great job! You can click [+ Batch] above to queue the next set of rows, or export your delivery report.';
    counter.textContent = `Completed (${range.length} / ${range.length})`;
    sendBtn.disabled = true;
    return;
  }

  const current = range[currentQueueIndex];
  const phone = current._phoneInfo;
  const tpl = messageTemplate.trim() || 'Hello! (Add your template in Step 3)';
  const msg = interpolate(tpl, current);

  const contactName = current[columns[0]] || `Contact ${currentQueueIndex + 1}`;
  const actualSrNo = (current._index !== undefined ? current._index + 1 : (currentQueueIndex + 1));

  counter.innerHTML = `<span style="color: var(--accent); font-weight: 800;">Sr No. #${actualSrNo}</span> of ${normalizedData.length} (Batch: ${currentQueueIndex + 1} / ${range.length})`;
  nameEl.textContent = contactName;
  phoneEl.innerHTML = phone.valid
    ? `📱 +${phone.clean} ${current._status === 'sent' ? '<span style="color: var(--success); font-size: 12px;">(Already Sent)</span>' : ''}`
    : `⚠️ <span style="color: var(--danger);">+${phone.clean || 'N/A'} (Invalid Phone)</span>`;
  msgEl.textContent = msg;
  sendBtn.disabled = false;

  updateProgressBar(currentQueueIndex, range.length);
  updateRangeStatus();
}

async function quickSendCurrent() {
  const range = getActiveRangeData();
  if (currentQueueIndex >= range.length) return;

  const current = range[currentQueueIndex];
  const phone = current._phoneInfo;
  const tpl = messageTemplate.trim() || 'Hello!';
  const msg = interpolate(tpl, current);
  const contactName = current[columns[0]] || `Contact ${currentQueueIndex + 1}`;

  if (!phone.valid) {
    const log = { phone: phone.clean || 'Invalid', name: contactName, message: msg };
    addLogEntry(log, 'skipped', phone.reason);
    current._status = 'skipped';
    if (activeCampaign) {
      await window.appDB.updateRowStatus(activeCampaign.id, current._index, 'skipped', log);
    }
    showToast(`Skipped invalid number for ${contactName}`, 'error');
  } else {
    const url = `https://wa.me/${phone.clean}?text=${encodeURIComponent(msg)}`;
    window.open(url, 'AetheriaWhatsAppTab');

    const log = { phone: phone.clean, name: contactName, message: msg };
    addLogEntry(log, 'sent', 'Opened in WhatsApp');
    current._status = 'sent';
    if (activeCampaign) {
      await window.appDB.updateRowStatus(activeCampaign.id, current._index, 'sent', log);
    }
    showToast(`✅ Opened for ${contactName}`, 'success');
  }

  renderDataTable();
  currentQueueIndex++;
  updateQuickQueueCard();
}

async function quickSkipCurrent() {
  const range = getActiveRangeData();
  if (currentQueueIndex >= range.length) return;

  const current = range[currentQueueIndex];
  const contactName = current[columns[0]] || `Contact ${currentQueueIndex + 1}`;

  const log = { phone: current._phoneInfo.clean || 'N/A', name: contactName, message: 'User skipped manually' };
  addLogEntry(log, 'skipped', 'Skipped manually');

  current._status = 'skipped';
  if (activeCampaign) {
    await window.appDB.updateRowStatus(activeCampaign.id, current._index, 'skipped', log);
  }

  renderDataTable();
  currentQueueIndex++;
  updateQuickQueueCard();
  showToast(`⏭️ Skipped ${contactName}`, 'info');
}

function quickResetQueue() {
  currentQueueIndex = 0;
  updateQuickQueueCard();
  showToast('🔄 Queue restarted from beginning', 'info');
}

// --- Mode B: Auto Loop ---
function startAutoDispatch() {
  if (!messageTemplate.trim()) {
    showToast('Please enter a message template in Step 3 first!', 'error');
    return;
  }

  const range = getActiveRangeData();
  if (range.length === 0) {
    showToast('No contacts in range to send.', 'error');
    return;
  }

  isAutoRunning = true;
  isAutoPaused = false;
  currentQueueIndex = 0;

  document.getElementById('autoStartBtn').style.display = 'none';
  document.getElementById('autoPauseBtn').style.display = 'inline-flex';
  document.getElementById('autoResumeBtn').style.display = 'none';

  showToast(`🚀 Auto loop started for ${range.length} contacts!`, 'success');
  autoSendNext();
}

async function autoSendNext() {
  const range = getActiveRangeData();

  if (!isAutoRunning || isAutoPaused || currentQueueIndex >= range.length) {
    if (currentQueueIndex >= range.length) {
      onAutoDispatchFinished();
    }
    return;
  }

  const current = range[currentQueueIndex];
  const phone = current._phoneInfo;
  const msg = interpolate(messageTemplate, current);
  const contactName = current[columns[0]] || `Contact ${currentQueueIndex + 1}`;

  if (phone.valid) {
    const url = `https://wa.me/${phone.clean}?text=${encodeURIComponent(msg)}`;
    window.open(url, 'AetheriaWhatsAppTab');

    const log = { phone: phone.clean, name: contactName, message: msg };
    addLogEntry(log, 'sent', 'Auto sent');
    current._status = 'sent';
    if (activeCampaign) {
      await window.appDB.updateRowStatus(activeCampaign.id, current._index, 'sent', log);
    }
  } else {
    const log = { phone: phone.clean || 'Invalid', name: contactName, message: msg };
    addLogEntry(log, 'skipped', phone.reason);
    current._status = 'skipped';
    if (activeCampaign) {
      await window.appDB.updateRowStatus(activeCampaign.id, current._index, 'skipped', log);
    }
  }

  renderDataTable();
  currentQueueIndex++;
  updateProgressBar(currentQueueIndex, range.length);

  if (currentQueueIndex < range.length && isAutoRunning && !isAutoPaused) {
    const delay = parseInt(document.getElementById('delaySlider').value) * 1000;
    autoTimer = setTimeout(() => autoSendNext(), delay);
  } else if (currentQueueIndex >= range.length) {
    onAutoDispatchFinished();
  }
}

function pauseAutoDispatch() {
  isAutoPaused = true;
  clearTimeout(autoTimer);
  document.getElementById('autoPauseBtn').style.display = 'none';
  document.getElementById('autoResumeBtn').style.display = 'inline-flex';
  showToast('⏸ Auto loop paused.', 'info');
}

function resumeAutoDispatch() {
  isAutoPaused = false;
  document.getElementById('autoPauseBtn').style.display = 'inline-flex';
  document.getElementById('autoResumeBtn').style.display = 'none';
  showToast('▶ Auto loop resumed!', 'success');
  autoSendNext();
}

function onAutoDispatchFinished() {
  isAutoRunning = false;
  isAutoPaused = false;
  clearTimeout(autoTimer);

  document.getElementById('autoStartBtn').style.display = 'inline-flex';
  document.getElementById('autoStartBtn').innerHTML = '<span class="btn-icon">🔄</span> Run Again';
  document.getElementById('autoPauseBtn').style.display = 'none';
  document.getElementById('autoResumeBtn').style.display = 'none';

  showToast('🎉 All automated messages processed and saved to database!', 'success');
}

// ========== PROGRESS & LOGS ==========
function updateProgressBar(done, total) {
  const percent = total > 0 ? Math.round((done / total) * 100) : 0;
  document.getElementById('progressBar').style.width = `${percent}%`;
  document.getElementById('progressText').textContent = `${done} / ${total} processed (${percent}%)`;
}

function addLogEntry(item, status, detail) {
  const logItem = { ...item, status, detail, timestamp: new Date().toLocaleTimeString() };
  dispatchLogs.push(logItem);

  const logList = document.getElementById('logList');
  const placeholder = logList.querySelector('.placeholder-text');
  if (placeholder) placeholder.remove();

  const statusIcons = {
    sent: '✅',
    failed: '❌',
    skipped: '⏭️',
    ready: '⏳'
  };

  const div = document.createElement('div');
  div.className = `log-item status-${status}`;
  div.innerHTML = `
    <span class="log-status">${statusIcons[status] || '•'}</span>
    <span class="log-phone">+${item.phone}</span>
    <span class="log-msg"><strong>${escapeHtml(item.name)}:</strong> ${escapeHtml(item.message).substring(0, 75)}...</span>
    <span class="log-time">${logItem.timestamp}</span>
  `;

  logList.prepend(div);
}

// ========== DATABASE OPERATIONS & PERSISTENCE ==========

async function saveCurrentCampaign(notifyUser = true) {
  if (parsedData.length === 0) {
    if (notifyUser) showToast('No data loaded to save.', 'error');
    return;
  }

  const campName = currentFileName || `Campaign-${new Date().toLocaleDateString()}`;
  const validCount = normalizedData.filter(r => r._phoneInfo && r._phoneInfo.valid).length;
  const sentCount = normalizedData.filter(r => r._status === 'sent').length;

  activeCampaign = {
    id: activeCampaign ? activeCampaign.id : 'camp_' + Date.now(),
    name: campName,
    fileName: currentFileName,
    createdAt: activeCampaign ? activeCampaign.createdAt : new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    totalContacts: parsedData.length,
    validContacts: validCount,
    phoneColumn: phoneColumn,
    countryCode: document.getElementById('countryCode').value,
    template: messageTemplate,
    columns: columns,
    data: normalizedData,
    logs: dispatchLogs,
    stats: {
      sent: sentCount,
      skipped: normalizedData.filter(r => r._status === 'skipped').length,
      failed: normalizedData.filter(r => r._status === 'failed').length
    }
  };

  await window.appDB.saveCampaign(activeCampaign);
  updateActiveCampaignBanner();
  await refreshDatabaseBadges();

  if (notifyUser) {
    showToast('💾 Campaign saved to local database!', 'success');
  }
}

function updateActiveCampaignBanner() {
  const banner = document.getElementById('activeCampaignNotice');
  if (!activeCampaign) {
    banner.style.display = 'none';
    return;
  }

  banner.style.display = 'flex';
  document.getElementById('activeCampName').textContent = activeCampaign.name;
  const sent = activeCampaign.stats ? activeCampaign.stats.sent || 0 : 0;
  document.getElementById('activeCampMeta').textContent =
    `${activeCampaign.totalContacts} contacts · ${sent} sent`;
}

function closeActiveCampaign() {
  activeCampaign = null;
  updateActiveCampaignBanner();
}

// --- History Table Renderer ---
async function renderHistoryTable() {
  const tbody = document.getElementById('historyTableBody');
  const campaigns = await window.appDB.getAllCampaigns();

  if (campaigns.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="5" class="empty-table-notice">
          No saved campaigns found in database. Upload an Excel or CSV file in the workspace to start!
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = '';
  campaigns.forEach(c => {
    const tr = document.createElement('tr');
    const createdDate = new Date(c.updatedAt || c.createdAt).toLocaleString();
    const sent = c.stats ? c.stats.sent || 0 : 0;
    const percent = c.totalContacts > 0 ? Math.round((sent / c.totalContacts) * 100) : 0;

    tr.innerHTML = `
      <td>
        <strong>${escapeHtml(c.name)}</strong>
        <div style="font-size: 11px; color: var(--text-muted);">${escapeHtml(c.fileName || 'Spreadsheet')}</div>
      </td>
      <td style="font-size: 12px; color: var(--text-secondary);">${createdDate}</td>
      <td><strong>${c.totalContacts}</strong> <span style="font-size: 11px; color: var(--accent);">(${c.validContacts} valid)</span></td>
      <td>
        <div class="progress-cell">
          <span>${sent} / ${c.totalContacts} (${percent}%)</span>
          <div class="mini-progress-bar"><div class="mini-progress-fill" style="width: ${percent}%;"></div></div>
        </div>
      </td>
      <td style="text-align: right;">
        <button class="btn-xs btn-primary" onclick="resumeCampaign('${c.id}')">📂 Open / Resume</button>
        <button class="btn-xs btn-secondary" onclick="exportCampaignReport('${c.id}')">📥 Report</button>
        <button class="btn-xs btn-danger" onclick="deleteCampaignFromDB('${c.id}')">🗑️</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

async function resumeCampaign(id) {
  const camp = await window.appDB.getCampaign(id);
  if (!camp) {
    showToast('Campaign not found', 'error');
    return;
  }

  activeCampaign = camp;
  currentFileName = camp.fileName || camp.name;
  parsedData = camp.data.map(d => {
    const raw = { ...d };
    delete raw._phoneInfo;
    delete raw._index;
    return raw;
  });
  columns = camp.columns || Object.keys(parsedData[0]);
  phoneColumn = camp.phoneColumn || columns[0];
  messageTemplate = camp.template || '';
  normalizedData = camp.data;

  // Sync inputs
  document.getElementById('countryCode').value = camp.countryCode || '91';
  document.getElementById('messageTemplate').value = messageTemplate;
  document.getElementById('fileName').textContent = camp.name;
  document.getElementById('fileSize').textContent = `(${camp.totalContacts} contacts)`;
  document.getElementById('fileInfo').style.display = 'block';
  document.getElementById('uploadZone').style.display = 'none';

  setupPhoneDropdown();
  setupRangeInputs();
  renderStatsBanner();
  renderDataTable();
  renderVariableTags();
  renderLivePreviews();
  updateQuickQueueCard();
  updateActiveCampaignBanner();

  document.getElementById('step2').style.display = 'block';
  document.getElementById('step3').style.display = 'block';
  document.getElementById('step4').style.display = 'block';

  switchView('workspace');
  showToast(`📂 Resumed campaign: "${camp.name}"!`, 'success');
}

async function deleteCampaignFromDB(id) {
  if (confirm('Are you sure you want to delete this campaign from history?')) {
    await window.appDB.deleteCampaign(id);
    if (activeCampaign && activeCampaign.id === id) {
      activeCampaign = null;
      updateActiveCampaignBanner();
    }
    await refreshDatabaseBadges();
    renderHistoryTable();
    showToast('Deleted campaign from database', 'info');
  }
}

async function exportCampaignReport(id) {
  const camp = await window.appDB.getCampaign(id);
  if (!camp) return;

  let csv = 'Sr No,Name,Phone,Status,Message\n';
  (camp.data || []).forEach((row, idx) => {
    const phone = row._phoneInfo ? row._phoneInfo.clean : '';
    const status = row._status || 'pending';
    const name = row[camp.columns[0]] || `Contact ${idx + 1}`;
    const msg = interpolate(camp.template || '', row);
    csv += `${idx + 1},"${name}","+${phone}","${status}","${msg.replace(/"/g, '""')}"\n`;
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${camp.name.replace(/[^a-zA-Z0-9_-]/g, '_')}-report.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('📥 Report downloaded!', 'success');
}

// ========== TEMPLATE LIBRARY ==========
async function populateTemplateDropdown() {
  const select = document.getElementById('templateQuickDropdown');
  select.innerHTML = '<option value="">📚 Load Saved Template...</option>';

  const templates = await window.appDB.getAllTemplates();
  templates.forEach(t => {
    const opt = document.createElement('option');
    opt.value = t.id;
    opt.textContent = t.title;
    select.appendChild(opt);
  });
}

async function loadQuickTemplate(id) {
  if (!id) return;
  const templates = await window.appDB.getAllTemplates();
  const found = templates.find(t => t.id === id);
  if (found) {
    messageTemplate = found.content;
    document.getElementById('messageTemplate').value = messageTemplate;
    renderLivePreviews();
    updateQuickQueueCard();
    showToast(`Loaded template: "${found.title}"`, 'info');
  }
}

async function promptSaveCurrentTemplate() {
  if (!messageTemplate.trim()) {
    showToast('Please write a message template first before saving!', 'error');
    return;
  }

  const title = prompt('Enter a title for this template:', 'Payment Reminder');
  if (!title) return;

  await window.appDB.saveTemplate({
    title: title.trim(),
    content: messageTemplate.trim()
  });

  await refreshDatabaseBadges();
  await populateTemplateDropdown();
  showToast(`💾 Template "${title}" saved to library!`, 'success');
}

async function promptCreateNewTemplate() {
  const title = prompt('Enter template title:');
  if (!title) return;
  const content = prompt('Enter template content (use {{Name}}, {{Amount}}, etc.):');
  if (!content) return;

  await window.appDB.saveTemplate({
    title: title.trim(),
    content: content.trim()
  });

  await refreshDatabaseBadges();
  await populateTemplateDropdown();
  renderTemplatesGrid();
  showToast(`✅ Created template "${title}"!`, 'success');
}

async function renderTemplatesGrid() {
  const grid = document.getElementById('templatesGrid');
  const templates = await window.appDB.getAllTemplates();

  if (templates.length === 0) {
    grid.innerHTML = '<p class="placeholder-text">No templates in library yet.</p>';
    return;
  }

  grid.innerHTML = '';
  templates.forEach(t => {
    const card = document.createElement('div');
    card.className = 'template-lib-card';
    card.innerHTML = `
      <div class="tpl-card-header">
        <h4>${escapeHtml(t.title)}</h4>
        <button class="btn-remove" onclick="deleteTemplateFromDB('${t.id}')" title="Delete template">✕</button>
      </div>
      <div class="tpl-card-body">${escapeHtml(t.content)}</div>
      <div class="tpl-card-footer">
        <button class="btn-primary btn-xs" onclick="useTemplateInWorkspace('${t.id}')">
          ⚡ Use This Template
        </button>
      </div>
    `;
    grid.appendChild(card);
  });
}

async function useTemplateInWorkspace(id) {
  await loadQuickTemplate(id);
  switchView('workspace');
  setTimeout(() => {
    const step3 = document.getElementById('step3');
    if (step3.style.display !== 'none') {
      step3.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, 100);
}

async function deleteTemplateFromDB(id) {
  if (confirm('Delete this template?')) {
    await window.appDB.deleteTemplate(id);
    await refreshDatabaseBadges();
    await populateTemplateDropdown();
    renderTemplatesGrid();
    showToast('Deleted template', 'info');
  }
}

// ========== BACKUP & RESTORE ==========
async function exportDatabaseBackup() {
  const backup = await window.appDB.exportFullBackup();
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Aetheria-AutoSend-Backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('📥 Full database backup downloaded!', 'success');
}

function triggerImportBackup() {
  document.getElementById('backupFileInput').click();
}

function handleImportBackupFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (evt) => {
    try {
      const data = JSON.parse(evt.target.result);
      await window.appDB.importFullBackup(data);
      await refreshDatabaseBadges();
      await populateTemplateDropdown();
      renderHistoryTable();
      showToast('🎉 Database restored successfully from backup!', 'success');
    } catch (err) {
      showToast('Error importing backup: ' + err.message, 'error');
    }
  };
  reader.readAsText(file);
}

// ========== EXPORT REPORT ==========
function exportReport() {
  if (activeCampaign && activeCampaign.id) {
    exportCampaignReport(activeCampaign.id);
  } else {
    showToast('No active campaign to export', 'error');
  }
}

// ========== RESET APPLICATION ==========
function resetApplication() {
  currentWorkbook = null;
  currentSheetName = '';
  currentFileName = '';
  activeCampaign = null;
  parsedData = [];
  normalizedData = [];
  columns = [];
  phoneColumn = '';
  messageTemplate = '';
  currentQueueIndex = 0;
  dispatchLogs.length = 0;
  isAutoRunning = false;
  isAutoPaused = false;
  clearTimeout(autoTimer);

  document.getElementById('uploadZone').style.display = 'block';
  document.getElementById('fileInfo').style.display = 'none';
  document.getElementById('sheetSelectorWrapper').style.display = 'none';
  document.getElementById('loadingBar').style.display = 'none';
  document.getElementById('fileInput').value = '';

  document.getElementById('step2').style.display = 'none';
  document.getElementById('step3').style.display = 'none';
  document.getElementById('step4').style.display = 'none';
  document.getElementById('messageTemplate').value = '';

  document.getElementById('logList').innerHTML = '<p class="placeholder-text">Messages will appear here as you process contacts...</p>';
  document.getElementById('progressBar').style.width = '0%';
  document.getElementById('progressText').textContent = '0 / 0 sent';

  updateActiveCampaignBanner();
}

// ========== UTILITIES ==========
function hideLoading() {
  document.getElementById('loadingBar').style.display = 'none';
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  const div = document.createElement('div');
  div.textContent = String(str);
  return div.innerHTML;
}

function showToast(message, type = 'success') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast toast-${type} show`;

  clearTimeout(toast._timeout);
  toast._timeout = setTimeout(() => {
    toast.classList.remove('show');
  }, 4000);
}
