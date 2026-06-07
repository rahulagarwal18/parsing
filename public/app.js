// Client ID for SSE streaming identification
const clientId = 'client-' + Math.random().toString(36).substring(2, 15);

// State Management
let appState = {
  currentStep: 1,
  excelFile: null,
  excelHeaders: [],
  excelPreview: [],
  selectedHeaders: [],
  matchKey: '',
  pdfFiles: [],
  comparisonReport: null
};

// DOM Elements
const stepNavs = [
  document.getElementById('step-nav-1'),
  document.getElementById('step-nav-2'),
  document.getElementById('step-nav-3'),
  document.getElementById('step-nav-4')
];

const panels = [
  document.getElementById('panel-step-1'),
  document.getElementById('panel-step-2'),
  document.getElementById('panel-step-3'),
  document.getElementById('panel-step-4')
];

// Initialise Lucide icons
document.addEventListener('DOMContentLoaded', () => {
  lucide.createIcons();
  setupEventListeners();
  initSSE();
});

// Setup SSE Progress Listener
let sseSource = null;
function initSSE() {
  const protocol = window.location.protocol;
  const host = window.location.host;
  sseSource = new EventSource(`/api/progress-stream?clientId=${clientId}`);

  sseSource.onmessage = (event) => {
    const data = JSON.parse(event.data);
    handleSSEMessage(data);
  };

  sseSource.onerror = (err) => {
    console.error("SSE Connection Error. Attempting reconnect...", err);
  };
}

// Event Listeners
function setupEventListeners() {
  // Step 1: Excel Upload Drag & Drop
  const excelDropZone = document.getElementById('excel-drop-zone');
  const excelInput = document.getElementById('excel-input');
  const removeExcelBtn = document.getElementById('remove-excel-btn');
  const btnGotoStep2 = document.getElementById('btn-goto-step2');

  excelDropZone.addEventListener('click', () => excelInput.click());
  excelDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    excelDropZone.classList.add('dragover');
  });
  excelDropZone.addEventListener('dragleave', () => excelDropZone.classList.remove('dragover'));
  excelDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    excelDropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      handleExcelFileSelect(e.dataTransfer.files[0]);
    }
  });

  excelInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleExcelFileSelect(e.target.files[0]);
    }
  });

  removeExcelBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    resetExcelUpload();
  });

  btnGotoStep2.addEventListener('click', () => {
    goToStep(2);
  });

  // Step 2: Config Panel
  const btnBackToStep1 = document.getElementById('btn-back-to-step1');
  const btnGotoStep3 = document.getElementById('btn-goto-step3');
  const matchKeySelect = document.getElementById('match-key-select');

  btnBackToStep1.addEventListener('click', () => goToStep(1));
  btnGotoStep3.addEventListener('click', () => goToStep(3));

  matchKeySelect.addEventListener('change', (e) => {
    appState.matchKey = e.target.value;
    validateStep2();
  });

  // Step 3: PDF Upload Drag & Drop
  const pdfDropZone = document.getElementById('pdf-drop-zone');
  const pdfInput = document.getElementById('pdf-input');
  const clearPdfsBtn = document.getElementById('clear-pdfs-btn');
  const btnBackToStep2 = document.getElementById('btn-back-to-step2');
  const btnRunPipeline = document.getElementById('btn-run-pipeline');

  pdfDropZone.addEventListener('click', () => pdfInput.click());
  pdfDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    pdfDropZone.classList.add('dragover');
  });
  pdfDropZone.addEventListener('dragleave', () => pdfDropZone.classList.remove('dragover'));
  pdfDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    pdfDropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      handlePdfFilesSelect(e.dataTransfer.files);
    }
  });

  pdfInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handlePdfFilesSelect(e.target.files);
    }
  });

  clearPdfsBtn.addEventListener('click', () => {
    appState.pdfFiles = [];
    updatePdfListUI();
  });

  btnBackToStep2.addEventListener('click', () => goToStep(2));
  btnRunPipeline.addEventListener('click', startPipeline);

  // Step 4: Report Results & Actions
  const btnRestartApp = document.getElementById('btn-restart-app');
  const searchInput = document.getElementById('result-search-input');
  const statusFilter = document.getElementById('status-filter-select');

  btnRestartApp.addEventListener('click', restartApplication);
  
  searchInput.addEventListener('input', filterReportTable);
  statusFilter.addEventListener('change', filterReportTable);
}

// Navigation flow helper
function goToStep(stepNumber) {
  appState.currentStep = stepNumber;

  // Update Stepper Nav
  stepNavs.forEach((nav, idx) => {
    const navStep = idx + 1;
    nav.classList.remove('active', 'completed');
    if (navStep === stepNumber) {
      nav.classList.add('active');
    } else if (navStep < stepNumber) {
      nav.classList.add('completed');
    }
  });

  // Update panels
  panels.forEach((panel, idx) => {
    panel.classList.remove('active');
    if (idx + 1 === stepNumber) {
      panel.classList.add('active');
    }
  });

  // Scroll to top of panel smoothly
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// --- STEP 1: EXCEL & PDF MASTER HANDLERS ---
function handleExcelFileSelect(file) {
  const isPdf = file.name.toLowerCase().endsWith('.pdf');
  if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls') && !isPdf) {
    alert("Please upload a valid Excel spreadsheet (.xlsx or .xls) or PDF document (.pdf)");
    return;
  }
  
  appState.excelFile = file;
  
  // Show file details
  document.getElementById('excel-file-name').innerText = file.name;
  document.getElementById('excel-file-size').innerText = formatBytes(file.size);
  document.getElementById('excel-drop-zone').style.display = 'none';
  document.getElementById('excel-file-banner').style.display = 'flex';

  // Upload to API immediately to extract headers and preview
  if (isPdf) {
    uploadPdfMasterToServer(file);
  } else {
    uploadExcelToServer(file);
  }
}

function resetExcelUpload() {
  appState.excelFile = null;
  appState.excelHeaders = [];
  appState.excelPreview = [];
  appState.selectedHeaders = [];
  appState.matchKey = '';

  document.getElementById('excel-drop-zone').style.display = 'flex';
  document.getElementById('excel-file-banner').style.display = 'none';
  document.getElementById('excel-preview-area').style.display = 'none';
  document.getElementById('btn-goto-step2').disabled = true;

  document.getElementById('excel-input').value = '';
}

async function uploadPdfMasterToServer(file) {
  const formData = new FormData();
  formData.append('pdfMasterFile', file);

  try {
    // Show spinner/progress or loading text in preview area
    document.getElementById('excel-preview-area').style.display = 'block';
    const previewTable = document.getElementById('excel-preview-table');
    previewTable.innerHTML = `<tr><td style="text-align:center; padding: 40px; color: var(--text-muted);">
      <div class="loading-spinner"></div>
      <div style="margin-top: 15px; font-weight: 500;">Parsing master PDF and extracting table records via LlamaParse & Gemini...</div>
      <div style="font-size: 11px; margin-top: 6px; opacity: 0.8;">This takes a moment as we digitize the PDF structure.</div>
    </td></tr>`;

    const res = await fetch('/api/upload-pdf-master', {
      method: 'POST',
      body: formData
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Failed to parse master PDF');
    }

    const data = await res.json();
    appState.excelHeaders = data.headers;
    appState.excelPreview = data.preview;

    // Populated header configuration layout (Step 2)
    populateStep2Config();
    
    // Render preview
    renderExcelPreviewTable();

    document.getElementById('btn-goto-step2').disabled = false;
  } catch (error) {
    alert("Error loading master PDF: " + error.message);
    resetExcelUpload();
  }
}

async function uploadExcelToServer(file) {
  const formData = new FormData();
  formData.append('excelFile', file);

  try {
    // Show spinner/progress or loading text in preview area
    document.getElementById('excel-preview-area').style.display = 'block';
    const previewTable = document.getElementById('excel-preview-table');
    previewTable.innerHTML = `<tr><td style="text-align:center; padding: 30px;">Reading spreadsheet structure...</td></tr>`;

    const res = await fetch('/api/upload-excel', {
      method: 'POST',
      body: formData
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Failed to upload spreadsheet');
    }

    const data = await res.json();
    appState.excelHeaders = data.headers;
    appState.excelPreview = data.preview;

    // Populated header configuration layout (Step 2)
    populateStep2Config();
    
    // Render Excel preview
    renderExcelPreviewTable();

    document.getElementById('btn-goto-step2').disabled = false;
  } catch (error) {
    alert("Error loading Excel file: " + error.message);
    resetExcelUpload();
  }
}

function renderExcelPreviewTable() {
  const table = document.getElementById('excel-preview-table');
  table.innerHTML = '';

  if (appState.excelHeaders.length === 0 || appState.excelPreview.length === 0) {
    return;
  }

  // Create header
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  appState.excelHeaders.forEach(h => {
    const th = document.createElement('th');
    th.innerText = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Create body
  const tbody = document.createElement('tbody');
  appState.excelPreview.forEach(row => {
    const tr = document.createElement('tr');
    appState.excelHeaders.forEach(h => {
      const td = document.createElement('td');
      td.innerText = row[h] !== undefined ? row[h] : '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

// --- STEP 2: CONFIGURATION HANDLERS ---
function populateStep2Config() {
  const headersSelector = document.getElementById('headers-selector');
  const matchKeySelect = document.getElementById('match-key-select');

  headersSelector.innerHTML = '';
  // Clear select dropdown except default template option
  matchKeySelect.innerHTML = '<option value="" disabled selected>Select unique matching key...</option>';

  // Default select all headers for comparison
  appState.selectedHeaders = [...appState.excelHeaders];

  appState.excelHeaders.forEach(header => {
    // 1. Create Badge item
    const badge = document.createElement('div');
    badge.className = 'field-badge selected';
    badge.dataset.header = header;
    badge.innerHTML = `<i data-lucide="check" class="badge-icon"></i> ${header}`;
    
    badge.addEventListener('click', () => {
      toggleHeaderSelection(badge, header);
    });

    headersSelector.appendChild(badge);

    // 2. Add as option to unique match key select dropdown
    const option = document.createElement('option');
    option.value = header;
    option.innerText = header;
    matchKeySelect.appendChild(option);
  });

  // Re-initialise Lucide icons inside badges
  lucide.createIcons();

  validateStep2();
}

function toggleHeaderSelection(badgeElement, header) {
  const idx = appState.selectedHeaders.indexOf(header);
  if (idx !== -1) {
    // Don't deselect the match key if already chosen
    if (header === appState.matchKey) {
      alert("You cannot deselect the matching key. Choose a different match key first.");
      return;
    }
    appState.selectedHeaders.splice(idx, 1);
    badgeElement.classList.remove('selected');
    badgeElement.querySelector('i').setAttribute('data-lucide', 'plus');
  } else {
    appState.selectedHeaders.push(header);
    badgeElement.classList.add('selected');
    badgeElement.querySelector('i').setAttribute('data-lucide', 'check');
  }

  // Update icons dynamically
  lucide.createIcons();
  
  // Update matching key list: only allow chosen comparison headers to be the matching key
  updateMatchKeyOptions();
  validateStep2();
}

function updateMatchKeyOptions() {
  const matchKeySelect = document.getElementById('match-key-select');
  const previousVal = appState.matchKey;
  
  matchKeySelect.innerHTML = '<option value="" disabled>Select unique matching key...</option>';
  
  appState.selectedHeaders.forEach(header => {
    const option = document.createElement('option');
    option.value = header;
    option.innerText = header;
    if (header === previousVal) {
      option.selected = true;
    }
    matchKeySelect.appendChild(option);
  });

  // If previous match key is no longer in selected headers, reset it
  if (!appState.selectedHeaders.includes(previousVal)) {
    appState.matchKey = '';
    matchKeySelect.value = '';
  }
}

function validateStep2() {
  const btnGotoStep3 = document.getElementById('btn-goto-step3');
  // Match key is mandatory and must be in selected headers list
  const isValid = appState.matchKey && appState.selectedHeaders.includes(appState.matchKey);
  btnGotoStep3.disabled = !isValid;
}

// --- STEP 3: PDF HANDLERS ---
function handlePdfFilesSelect(files) {
  const validPdfs = Array.from(files).filter(f => f.name.endsWith('.pdf'));

  if (validPdfs.length === 0) {
    alert("Please select valid PDF documents.");
    return;
  }

  // Ensure total files does not exceed limit
  const currentCount = appState.pdfFiles.length;
  const newCount = validPdfs.length;

  if (currentCount + newCount > 200) {
    alert("Limit exceeded: You can upload a maximum of 200 PDFs.");
    return;
  }

  // Add files to state (de-duplicate by name if necessary)
  validPdfs.forEach(file => {
    if (!appState.pdfFiles.some(f => f.name === file.name)) {
      appState.pdfFiles.push(file);
    }
  });

  updatePdfListUI();
}

function updatePdfListUI() {
  const gridContainer = document.getElementById('pdf-file-grid-container');
  const pdfListGrid = document.getElementById('pdf-list-grid');
  const countLabel = document.getElementById('pdf-count-label');
  const btnRunPipeline = document.getElementById('btn-run-pipeline');

  pdfListGrid.innerHTML = '';

  if (appState.pdfFiles.length === 0) {
    gridContainer.style.display = 'none';
    btnRunPipeline.disabled = true;
    return;
  }

  gridContainer.style.display = 'block';
  countLabel.innerText = `${appState.pdfFiles.length} PDF(s) Selected`;
  btnRunPipeline.disabled = false;

  appState.pdfFiles.forEach((file, index) => {
    const card = document.createElement('div');
    card.className = 'pdf-file-card';
    card.innerHTML = `
      <div class="pdf-card-info">
        <i data-lucide="file" class="pdf-card-icon"></i>
        <span class="pdf-card-name" title="${file.name}">${file.name}</span>
      </div>
      <button class="remove-file-btn"><i data-lucide="x"></i></button>
    `;

    card.querySelector('button').addEventListener('click', () => {
      appState.pdfFiles.splice(index, 1);
      updatePdfListUI();
    });

    pdfListGrid.appendChild(card);
  });

  lucide.createIcons();
}

// --- PIPELINE RUN AND SSE PROCESSOR ---
async function startPipeline() {
  // Hide Uploader form, reveal logs
  document.getElementById('pdf-upload-group').style.display = 'none';
  document.getElementById('pipeline-progress').style.display = 'block';
  
  // Block going back
  document.getElementById('step-nav-1').style.pointerEvents = 'none';
  document.getElementById('step-nav-2').style.pointerEvents = 'none';
  document.getElementById('step-nav-3').style.pointerEvents = 'none';

  // Clear log console
  const consoleLogs = document.getElementById('console-logs');
  consoleLogs.innerHTML = `<div class="log-line system">[SYSTEM] Starting batch reconciliation task...</div>`;

  // Update Status Indicators
  document.querySelector('.status-indicator').className = 'status-indicator busy';
  document.querySelector('.status-text').innerText = 'Pipeline Status: Processing';

  // Prepare Multipart Request Payload
  const formData = new FormData();
  formData.append('clientId', clientId);
  formData.append('matchKey', appState.matchKey);
  formData.append('selectedHeaders', JSON.stringify(appState.selectedHeaders));
  
  const customPrompt = document.getElementById('custom-prompt-input').value;
  formData.append('customPrompt', customPrompt);
  
  appState.pdfFiles.forEach(file => {
    formData.append('pdfFiles', file);
  });

  try {
    const res = await fetch('/api/process-pdfs', {
      method: 'POST',
      body: formData
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Failed to start parsing engine');
    }

    // Server responds with { success: true } immediately, processing continues in background.
    // SSE event stream handles logging and final reconciliation callback.

  } catch (error) {
    writeLog(`[ERROR] Processing initiation failed: ${error.message}`, 'failed');
    
    // Restore layout
    document.getElementById('pdf-upload-group').style.display = 'flex';
    document.getElementById('pipeline-progress').style.display = 'none';
    document.querySelector('.status-indicator').className = 'status-indicator online';
    document.querySelector('.status-text').innerText = 'Pipeline Status: Ready';
    
    alert("Reconciliation Pipeline Error: " + error.message);
  }
}

function handleSSEMessage(data) {
  const percentLabel = document.getElementById('progress-percent');
  const fillBar = document.getElementById('progress-bar-fill');
  const statusTitle = document.getElementById('progress-status-title');
  const statusSubtitle = document.getElementById('progress-status-subtitle');

  if (data.status === "start") {
    writeLog(`[SYSTEM] ${data.message}`, 'system');
    statusTitle.innerText = "Reconciliation Started";
    statusSubtitle.innerText = data.message;
    percentLabel.innerText = "0%";
    fillBar.style.width = "0%";
  }

  if (data.status === "parsing") {
    writeLog(`[PARSING] File: ${data.filename}`, 'parsing');
    statusTitle.innerText = "Parsing PDFs";
    statusSubtitle.innerText = `LlamaParse processing: ${data.filename}`;
    
    const percent = Math.round((data.index / appState.pdfFiles.length) * 100);
    percentLabel.innerText = `${percent}%`;
    fillBar.style.width = `${percent}%`;
  }

  if (data.status === "extracting") {
    writeLog(`[EXTRACT] Querying Gemini for fields in: ${data.filename}`, 'extracting');
    statusTitle.innerText = "Extracting Fields";
    statusSubtitle.innerText = `Gemini structured analysis: ${data.filename}`;
  }

  if (data.status === "completed_file") {
    writeLog(`[SUCCESS] Completed file reconciliation: ${data.filename}`, 'completed');
  }

  if (data.status === "failed_file") {
    writeLog(`[FAILED] Processing issue: ${data.message}`, 'failed');
  }

  if (data.status === "comparing") {
    writeLog(`[COMPARING] ${data.message}`, 'system');
    statusTitle.innerText = "Running Reconciliation Engine";
    statusSubtitle.innerText = data.message;
  }

  if (data.status === "done") {
    writeLog(`[SYSTEM] ${data.message}`, 'completed');
    
    percentLabel.innerText = "100%";
    fillBar.style.width = "100%";

    // Save final report data
    appState.comparisonReport = data.data;

    // Transition to Step 4 after brief delay to let user see 100%
    setTimeout(() => {
      renderStep4Dashboard();
      goToStep(4);
      
      // Reset Status
      document.querySelector('.status-indicator').className = 'status-indicator online';
      document.querySelector('.status-text').innerText = 'Pipeline Status: Ready';
    }, 1000);
  }

  if (data.status === "error") {
    writeLog(`[ERROR] ${data.message}`, 'failed');
    alert(`Audit Pipeline Error: ${data.message}`);
    
    // Restore layouts
    document.getElementById('pdf-upload-group').style.display = 'flex';
    document.getElementById('pipeline-progress').style.display = 'none';
    document.querySelector('.status-indicator').className = 'status-indicator online';
    document.querySelector('.status-text').innerText = 'Pipeline Status: Ready';
  }
}

function writeLog(message, type) {
  const consoleLogs = document.getElementById('console-logs');
  const line = document.createElement('div');
  const timestamp = new Date().toLocaleTimeString();
  line.className = `log-line ${type}`;
  line.innerText = `[${timestamp}] ${message}`;
  consoleLogs.appendChild(line);
  
  // Auto scroll console
  consoleLogs.scrollTop = consoleLogs.scrollHeight;
}

// --- STEP 4: REPORTING AND COMPARISON RESULTS ---
function renderStep4Dashboard() {
  const report = appState.comparisonReport;
  if (!report) return;

  // 1. Stat cards updates
  document.getElementById('stat-total-processed').innerText = report.summary.totalPdfsProcessed;
  document.getElementById('stat-matches').innerText = report.summary.totalMatches;
  document.getElementById('stat-mismatches').innerText = report.summary.totalMismatches;
  document.getElementById('stat-no-match').innerText = report.summary.totalNoExcelMatch;

  // 2. Configure export links
  document.getElementById('export-excel-btn').href = `/api/export/excel`;
  document.getElementById('export-pdf-btn').href = `/api/export/pdf`;
  document.getElementById('export-word-btn').href = `/api/export/word`;

  // 3. Render Detail Log Table
  renderReportTable();

  // 4. Render Unmatched Excel Table
  renderUnmatchedExcelTable();
}

function renderReportTable() {
  const table = document.getElementById('report-detail-table');
  const report = appState.comparisonReport;
  
  table.innerHTML = '';

  if (!report || report.records.length === 0) {
    table.innerHTML = '<tbody><tr><td style="text-align:center;">No records compared.</td></tr></tbody>';
    return;
  }

  // Create Header Row
  const thead = document.createElement('thead');
  const headerTr = document.createElement('tr');
  headerTr.innerHTML = `
    <th style="width: 40px;"></th>
    <th>PDF Document Name</th>
    <th>Status</th>
    <th>Matched Excel Row</th>
    <th>${appState.matchKey} Value</th>
  `;
  thead.appendChild(headerTr);
  table.appendChild(thead);

  // Create Body Rows
  const tbody = document.createElement('tbody');
  
  report.records.forEach((record, index) => {
    // Main row
    const tr = document.createElement('tr');
    tr.className = 'main-row';
    tr.dataset.index = index;

    let statusClass = '';
    let statusLabel = '';
    if (record.status === 'FULL_MATCH') {
      statusClass = 'match';
      statusLabel = 'Match';
    } else if (record.status === 'MISMATCH') {
      statusClass = 'mismatch';
      statusLabel = 'Mismatch';
    } else if (record.status === 'PARTIAL_MATCH') {
      statusClass = 'partial';
      statusLabel = 'Partial';
    } else {
      statusClass = 'nomatch';
      statusLabel = 'Unmatched';
    }

    const keyVal = record.matchedExcelRow ? record.matchedExcelRow[appState.matchKey] || 'N/A' : 'N/A';

    tr.innerHTML = `
      <td><i data-lucide="chevron-right" class="arrow-toggle" id="arrow-${index}"></i></td>
      <td style="font-weight: 600;">${record.filename}</td>
      <td><span class="status-badge ${statusClass}">${statusLabel}</span></td>
      <td>${record.excelRowIndex ? `Row ${record.excelRowIndex}` : '<span style="color:var(--text-muted)">N/A</span>'}</td>
      <td>${keyVal}</td>
    `;

    tbody.appendChild(tr);

    // Dynamic expanded row structure
    const expTr = document.createElement('tr');
    expTr.className = 'expanded-row';
    expTr.id = `expanded-row-${index}`;
    expTr.style.display = 'none';

    // Build the grid contents comparing each field
    let fieldsHtml = '';
    appState.selectedHeaders.forEach(header => {
      const field = record.fields[header] || { excelValue: 'N/A', pdfValue: '', status: 'MISSING_IN_PDF' };
      
      let cardClass = 'matched';
      let statusText = 'MATCHED';
      
      if (field.status === 'MISMATCH') {
        cardClass = 'mismatched';
        statusText = 'DISCREPANCY';
      } else if (field.status === 'MISSING_IN_PDF') {
        cardClass = 'missing';
        statusText = 'MISSING IN PDF';
      } else if (field.status === 'NO_EXCEL_MATCH') {
        cardClass = 'missing';
        statusText = 'NO MASTER VALUE';
      }

      fieldsHtml += `
        <div class="detail-field-card ${cardClass}">
          <div style="display:flex; justify-content:space-between; align-items:center;">
            <span class="field-name-label">${header}</span>
            <span class="field-status-badge ${cardClass}">${statusText}</span>
          </div>
          <div class="comparison-values">
            <div class="val-box">
              <span class="val-label">Excel Value</span>
              <span class="val-text">${field.excelValue !== undefined && field.excelValue !== '' ? field.excelValue : '<em style="color:var(--text-muted)">empty</em>'}</span>
            </div>
            <div class="val-box">
              <span class="val-label">PDF Extracted</span>
              <span class="val-text">${field.pdfValue !== undefined && field.pdfValue !== '' ? field.pdfValue : '<em style="color:var(--text-muted)">empty</em>'}</span>
            </div>
          </div>
        </div>
      `;
    });

    expTr.innerHTML = `
      <td colspan="5">
        <div class="expansion-details-container">
          <div class="expanded-title">Field Comparison breakdown</div>
          <div class="details-grid">
            ${fieldsHtml}
          </div>
        </div>
      </td>
    `;

    tbody.appendChild(expTr);

    // Click handler for expansion
    tr.querySelector('.arrow-toggle').addEventListener('click', (e) => {
      e.stopPropagation();
      toggleRowExpansion(index);
    });

    tr.addEventListener('click', () => {
      toggleRowExpansion(index);
    });
  });

  table.appendChild(tbody);
  lucide.createIcons();
}

function toggleRowExpansion(index) {
  const expRow = document.getElementById(`expanded-row-${index}`);
  const arrow = document.getElementById(`arrow-${index}`);
  
  if (expRow.style.display === 'none') {
    expRow.style.display = 'table-row';
    arrow.classList.add('open');
  } else {
    expRow.style.display = 'none';
    arrow.classList.remove('open');
  }
}

function renderUnmatchedExcelTable() {
  const table = document.getElementById('unmatched-excel-table');
  const report = appState.comparisonReport;

  table.innerHTML = '';

  if (!report || report.unmatchedExcelRows.length === 0) {
    table.innerHTML = '<tbody><tr><td style="text-align:center; padding: 15px;">All Excel rows were matched with PDF files!</td></tr></tbody>';
    return;
  }

  // Create headers
  const thead = document.createElement('thead');
  const headerTr = document.createElement('tr');
  headerTr.innerHTML = `<th>Row Number</th>`;
  
  // Get keys from first item
  const sampleRow = report.unmatchedExcelRows[0].data;
  const keys = Object.keys(sampleRow);

  keys.forEach(k => {
    const th = document.createElement('th');
    th.innerText = k;
    headerTr.appendChild(th);
  });
  thead.appendChild(headerTr);
  table.appendChild(thead);

  // Create body
  const tbody = document.createElement('tbody');
  report.unmatchedExcelRows.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td style="font-weight:600;">Row ${item.excelRowIndex}</td>`;
    keys.forEach(k => {
      const td = document.createElement('td');
      td.innerText = item.data[k] !== undefined ? item.data[k] : '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

// Client-side search and status filters
function filterReportTable() {
  const query = document.getElementById('result-search-input').value.toLowerCase();
  const filter = document.getElementById('status-filter-select').value;
  const records = appState.comparisonReport?.records || [];

  const mainRows = document.querySelectorAll('#report-detail-table tbody tr.main-row');
  
  mainRows.forEach(row => {
    const index = parseInt(row.dataset.index);
    const record = records[index];
    const expRow = document.getElementById(`expanded-row-${index}`);
    const arrow = document.getElementById(`arrow-${index}`);
    
    if (!record) return;

    const matchesSearch = record.filename.toLowerCase().includes(query) || 
                         (record.matchedExcelRow && String(record.matchedExcelRow[appState.matchKey]).toLowerCase().includes(query));
    
    let matchesStatus = true;
    if (filter !== 'ALL') {
      matchesStatus = record.status === filter;
    }

    if (matchesSearch && matchesStatus) {
      row.style.display = 'table-row';
      // keep expansion closed on filter updates
      expRow.style.display = 'none';
      arrow.classList.remove('open');
    } else {
      row.style.display = 'none';
      expRow.style.display = 'none';
      arrow.classList.remove('open');
    }
  });
}

// Restart Application
function restartApplication() {
  if (confirm("Are you sure you want to start a new reconciliation project? Current results will be reset.")) {
    // Reset state
    appState = {
      currentStep: 1,
      excelFile: null,
      excelHeaders: [],
      excelPreview: [],
      selectedHeaders: [],
      matchKey: '',
      pdfFiles: [],
      comparisonReport: null
    };

    // Reset UI forms
    resetExcelUpload();
    updatePdfListUI();
    document.getElementById('custom-prompt-input').value = '';
    
    // Enable back going
    document.getElementById('step-nav-1').style.pointerEvents = 'auto';
    document.getElementById('step-nav-2').style.pointerEvents = 'auto';
    document.getElementById('step-nav-3').style.pointerEvents = 'auto';

    // Show step 1
    document.getElementById('pdf-upload-group').style.display = 'flex';
    document.getElementById('pipeline-progress').style.display = 'none';
    
    goToStep(1);
  }
}

// --- UTILITY METHODS ---
function formatBytes(bytes, decimals = 2) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}
