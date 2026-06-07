// Client ID for SSE streaming identification
const clientId = 'client-' + Math.random().toString(36).substring(2, 15);

// State Management
let appState = {
  currentStep: 1,
  pdfFiles: [],
  pdfHeaders: [],
  pdfRows: [],
  excelFile: null,
  excelHeaders: [],
  excelPreview: [],
  mappings: {}, // { [pdfHeader]: excelHeader }
  matchKey: '', // pdfHeader
  comparisonReport: null,
  executiveSummary: ''
};

let currentPdfBlobUrl = null;

function updateSystemStatus(className, text) {
  const indicator = document.querySelector('.status-indicator');
  const textEl = document.querySelector('.status-text');
  if (indicator) indicator.className = `status-indicator ${className}`;
  if (textEl) textEl.innerText = text;
}

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

// Initialise Lucide icons and app controllers
document.addEventListener('DOMContentLoaded', () => {
  lucide.createIcons();
  setupEventListeners();
  initSSE();
});

// Setup SSE Progress Listener
let sseSource = null;
function initSSE() {
  sseSource = new EventSource(`/api/progress-stream?clientId=${clientId}`);

  sseSource.onmessage = (event) => {
    const data = JSON.parse(event.data);
    handleSSEMessage(data);
  };

  sseSource.onerror = (err) => {
    console.error("SSE Connection Error. Attempting reconnect...", err);
  };
}

// Event Listeners setup
function setupEventListeners() {
  // --- STEP 1: PDF Upload & Parsing ---
  const pdfDropZone = document.getElementById('pdf-drop-zone');
  const pdfInput = document.getElementById('pdf-input');
  const clearPdfsBtn = document.getElementById('clear-pdfs-btn');
  const btnRunPipeline = document.getElementById('btn-run-pipeline');
  const btnGotoStep2 = document.getElementById('btn-goto-step2');

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

  btnRunPipeline.addEventListener('click', startPdfParsing);
  btnGotoStep2.addEventListener('click', () => goToStep(2));

  // --- STEP 2: Parsed PDF Preview ---
  const btnBackToStep1 = document.getElementById('btn-back-to-step1');
  const btnGotoStep3 = document.getElementById('btn-goto-step3');

  btnBackToStep1.addEventListener('click', () => {
    // Confirm go back
    if (confirm("Are you sure you want to go back? Current parsed PDF data will be reset.")) {
      resetPdfParsing();
      goToStep(1);
    }
  });

  btnGotoStep3.addEventListener('click', () => {
    goToStep(3);
  });

  // --- STEP 3: Excel Upload & Mapping ---
  const excelDropZone = document.getElementById('excel-drop-zone');
  const excelInput = document.getElementById('excel-input');
  const removeExcelBtn = document.getElementById('remove-excel-btn');
  const btnBackToStep2 = document.getElementById('btn-back-to-step2');
  const btnAiAutoMap = document.getElementById('btn-ai-auto-map');
  const matchKeySelect = document.getElementById('match-key-select');
  const btnRunReconciliation = document.getElementById('btn-run-reconciliation');

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

  btnBackToStep2.addEventListener('click', () => goToStep(2));

  btnAiAutoMap.addEventListener('click', runAiAutoMapping);

  matchKeySelect.addEventListener('change', (e) => {
    appState.matchKey = e.target.value;
    validateReconciliationTrigger();
  });

  btnRunReconciliation.addEventListener('click', runReconciliation);

  // --- STEP 4: Reporting & Reset ---
  const btnRestartApp = document.getElementById('btn-restart-app');
  const searchInput = document.getElementById('result-search-input');
  const statusFilter = document.getElementById('status-filter-select');

  btnRestartApp.addEventListener('click', restartApplication);
  searchInput.addEventListener('input', filterReportTable);
  statusFilter.addEventListener('change', filterReportTable);

  // --- MODAL: Parsed Text View ---
  const closeModalBtn = document.getElementById('close-modal-btn');
  const modalOverlay = document.getElementById('parsed-text-modal');
  if (closeModalBtn && modalOverlay) {
    const closeModal = () => {
      modalOverlay.style.display = 'none';
      const container = document.getElementById('modal-pdf-viewer-container');
      if (container) container.innerHTML = '';
      if (currentPdfBlobUrl) {
        URL.revokeObjectURL(currentPdfBlobUrl);
        currentPdfBlobUrl = null;
      }
    };
    closeModalBtn.addEventListener('click', closeModal);
    modalOverlay.addEventListener('click', (e) => { if (e.target === modalOverlay) closeModal(); });
  }

  // --- MODAL TABS: switcher logic ---
  const tabs = document.querySelectorAll('.pane-tab');
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      // Deactivate all tabs
      tabs.forEach(t => {
        t.classList.remove('active');
        t.style.color = 'var(--text-muted)';
        t.style.borderBottomColor = 'transparent';
        t.style.fontWeight = '500';
      });

      // Activate clicked tab
      tab.classList.add('active');
      tab.style.color = '#fff';
      tab.style.borderBottomColor = 'var(--cyan)';
      tab.style.fontWeight = '600';

      // Hide all content panels
      const panels = document.querySelectorAll('.tab-content-panel');
      panels.forEach(p => p.style.display = 'none');

      // Show the selected panel
      const targetTab = tab.dataset.tab;
      const targetPanel = document.getElementById(`modal-tab-${targetTab}-content`);
      if (targetPanel) {
        targetPanel.style.display = 'block';
      }
    });
  });
}

// Stepper navigation flow helper
function goToStep(stepNumber) {
  appState.currentStep = stepNumber;

  // Update Stepper Navigation Status
  stepNavs.forEach((nav, idx) => {
    const navStep = idx + 1;
    nav.classList.remove('active', 'completed');
    if (navStep === stepNumber) {
      nav.classList.add('active');
    } else if (navStep < stepNumber) {
      nav.classList.add('completed');
    }
  });

  // Toggle active card panels
  panels.forEach((panel, idx) => {
    panel.classList.remove('active');
    if (idx + 1 === stepNumber) {
      panel.classList.add('active');
    }
  });

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// --- STEP 1: PDF BATCH HANDLERS ---
function handlePdfFilesSelect(files) {
  const validPdfs = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.pdf'));

  if (validPdfs.length === 0) {
    alert("Please select valid PDF documents.");
    return;
  }

  const currentCount = appState.pdfFiles.length;
  const newCount = validPdfs.length;

  if (currentCount + newCount > 200) {
    alert("Limit exceeded: You can upload a maximum of 200 PDFs.");
    return;
  }

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

    card.querySelector('button').addEventListener('click', (e) => {
      e.stopPropagation();
      appState.pdfFiles.splice(index, 1);
      updatePdfListUI();
    });

    pdfListGrid.appendChild(card);
  });

  lucide.createIcons();
}

function resetPdfParsing() {
  appState.pdfFiles = [];
  appState.pdfHeaders = [];
  appState.pdfRows = [];
  
  document.getElementById('pdf-upload-group').style.display = 'flex';
  document.getElementById('pipeline-progress').style.display = 'none';
  document.getElementById('btn-goto-step2').style.display = 'none';
  updatePdfListUI();
}

async function startPdfParsing() {
  document.getElementById('pdf-upload-group').style.display = 'none';
  document.getElementById('pipeline-progress').style.display = 'block';
  
  // Disable stepper navigation clicking during task execution
  document.getElementById('step-nav-1').style.pointerEvents = 'none';
  document.getElementById('step-nav-2').style.pointerEvents = 'none';
  document.getElementById('step-nav-3').style.pointerEvents = 'none';

  const progressList = document.getElementById('file-progress-list');
  if (progressList) {
    progressList.innerHTML = '';
    
    // Initialize file progress list visually
    appState.pdfFiles.forEach((file, idx) => {
      const card = document.createElement('div');
      card.id = `file-progress-card-${idx}`;
      card.className = 'file-progress-card';
      card.style.background = 'rgba(255, 255, 255, 0.01)';
      card.style.border = '1px solid rgba(255, 255, 255, 0.04)';
      card.style.borderRadius = '8px';
      card.style.padding = '12px 15px';
      card.style.display = 'flex';
      card.style.alignItems = 'center';
      card.style.justifyContent = 'space-between';
      card.style.gap = '10px';
      card.innerHTML = `
        <div style="display: flex; align-items: center; gap: 10px; overflow: hidden; flex: 1;">
          <i data-lucide="file-text" style="width: 16px; height: 16px; color: var(--text-muted); flex-shrink: 0;"></i>
          <span style="font-size: 13px; font-weight: 500; color: #fff; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${file.name}">${file.name}</span>
        </div>
        <div class="file-status-badge" style="font-size: 12px; font-weight: 600; color: var(--text-muted); display: flex; align-items: center; gap: 6px;">
          <i data-lucide="clock" style="width: 14px; height: 14px;"></i> Queued
        </div>
      `;
      progressList.appendChild(card);
    });
    lucide.createIcons();
  }

  const queueStatus = document.getElementById('queue-status-text');
  if (queueStatus) {
    queueStatus.innerText = `0 of ${appState.pdfFiles.length} files parsed`;
    queueStatus.style.color = 'var(--cyan)';
  }

  // Update Status Indicators safely
  updateSystemStatus('busy', 'Pipeline Status: Processing');

  const formData = new FormData();
  formData.append('clientId', clientId);
  appState.pdfFiles.forEach(file => {
    formData.append('pdfFiles', file);
  });

  try {
    const res = await fetch('/api/upload-pdf-batch', {
      method: 'POST',
      body: formData
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Failed to start parsing engine');
    }
  } catch (error) {
    writeLog(`[ERROR] Parsing task initiation failed: ${error.message}`, 'failed');
    
    // Restore layout
    document.getElementById('pdf-upload-group').style.display = 'flex';
    document.getElementById('pipeline-progress').style.display = 'none';
    updateSystemStatus('online', 'Pipeline Status: Ready');
    
    alert("PDF Parsing Error: " + error.message);
  }
}

function handleSSEMessage(data) {
  const percentLabel = document.getElementById('progress-percent');
  const fillBar = document.getElementById('progress-bar-fill');
  const statusTitle = document.getElementById('progress-status-title');
  const statusSubtitle = document.getElementById('progress-status-subtitle');

  if (data.status === "start") {
    writeLog(`[SYSTEM] ${data.message}`, 'system');
    statusTitle.innerText = "PDF Parsing Started";
    statusSubtitle.innerText = data.message;
    percentLabel.innerText = "0%";
    fillBar.style.width = "0%";
  }

  if (data.status === "parsing") {
    writeLog(`[PARSING] File: ${data.filename}`, 'parsing');
    statusTitle.innerText = "LlamaParse Digiting";
    statusSubtitle.innerText = `Processing layout: ${data.filename}`;
    
    const percent = Math.round((data.index / appState.pdfFiles.length) * 100);
    percentLabel.innerText = `${percent}%`;
    fillBar.style.width = `${percent}%`;

    // Update individual file card status
    const card = document.getElementById(`file-progress-card-${data.index}`);
    if (card) {
      card.style.borderColor = 'rgba(56, 189, 248, 0.3)';
      card.style.background = 'rgba(56, 189, 248, 0.02)';
      const badge = card.querySelector('.file-status-badge');
      if (badge) {
        badge.style.color = 'var(--cyan)';
        badge.innerHTML = `<div class="loading-spinner" style="width: 12px; height: 12px; border-width: 1.5px; border-top-color: var(--cyan); margin-right: 2px;"></div> Parsing`;
      }
    }
    const queueStatus = document.getElementById('queue-status-text');
    if (queueStatus) {
      queueStatus.innerText = `Processing file ${data.index + 1} of ${appState.pdfFiles.length}`;
    }
  }

  if (data.status === "detecting_schema") {
    writeLog(`[AI SCHEMA] Analyzing document structure in: ${data.filename}`, 'extracting');
    statusTitle.innerText = "Analyzing Schema";
    statusSubtitle.innerText = "Detecting common column fields from document...";
  }

  if (data.status === "schema_detected") {
    writeLog(`[AI SCHEMA] Common schema headers: ${data.headers.join(", ")}`, 'system');
  }

  if (data.status === "extracting") {
    writeLog(`[EXTRACT] Claude structured mapping: ${data.filename}`, 'extracting');
    statusTitle.innerText = "Extracting Fields";
    statusSubtitle.innerText = `Claude extracting headers: ${data.filename}`;

    // Update individual file card status
    const card = document.getElementById(`file-progress-card-${data.index}`);
    if (card) {
      card.style.borderColor = 'rgba(245, 158, 11, 0.3)';
      card.style.background = 'rgba(245, 158, 11, 0.02)';
      const badge = card.querySelector('.file-status-badge');
      if (badge) {
        badge.style.color = 'var(--warning)';
        badge.innerHTML = `<div class="loading-spinner" style="width: 12px; height: 12px; border-width: 1.5px; border-top-color: var(--warning); margin-right: 2px;"></div> Extracting`;
      }
    }
  }

  if (data.status === "completed_file") {
    writeLog(`[SUCCESS] Completed parsing: ${data.filename}`, 'completed');

    // Update individual file card status
    const card = document.getElementById(`file-progress-card-${data.index}`);
    if (card) {
      card.style.borderColor = 'rgba(16, 185, 129, 0.2)';
      card.style.background = 'rgba(16, 185, 129, 0.02)';
      const badge = card.querySelector('.file-status-badge');
      if (badge) {
        badge.style.color = 'var(--success)';
        badge.innerHTML = `<i data-lucide="check-circle" style="width: 14px; height: 14px; color: var(--success);"></i> Done`;
        lucide.createIcons();
      }
    }
    const queueStatus = document.getElementById('queue-status-text');
    if (queueStatus) {
      queueStatus.innerText = `${data.index + 1} of ${appState.pdfFiles.length} files parsed`;
    }
  }

  if (data.status === "failed_file") {
    writeLog(`[FAILED] File issue: ${data.message}`, 'failed');

    // Update individual file card status
    const card = document.getElementById(`file-progress-card-${data.index}`);
    if (card) {
      card.style.borderColor = 'rgba(239, 68, 68, 0.2)';
      card.style.background = 'rgba(239, 68, 68, 0.02)';
      const badge = card.querySelector('.file-status-badge');
      if (badge) {
        badge.style.color = 'var(--danger)';
        badge.innerHTML = `<i data-lucide="x-circle" style="width: 14px; height: 14px; color: var(--danger);"></i> Failed`;
        lucide.createIcons();
      }
    }
  }

  if (data.status === "done") {
    writeLog(`[SYSTEM] ${data.message}`, 'completed');
    
    percentLabel.innerText = "100%";
    fillBar.style.width = "100%";

    // Cache PDF extracted records
    appState.pdfHeaders = data.data.headers;
    appState.pdfRows = data.data.rows;

    // Render Preview Table in Step 2
    renderParsedPdfTable();

    // Show Continue Button
    document.getElementById('btn-goto-step2').style.display = 'flex';
    statusTitle.innerText = "Parsing Complete";
    statusSubtitle.innerText = "All PDF data parsed successfully.";

    // Reset status badge
    updateSystemStatus('online', 'Pipeline Status: Ready');
    
    // Enable stepper clicks
    document.getElementById('step-nav-1').style.pointerEvents = 'auto';
    document.getElementById('step-nav-2').style.pointerEvents = 'auto';

    const queueStatus = document.getElementById('queue-status-text');
    if (queueStatus) {
      queueStatus.innerText = `Completed`;
      queueStatus.style.color = 'var(--success)';
    }
  }

  if (data.status === "error") {
    writeLog(`[ERROR] ${data.message}`, 'failed');
    alert(`Reconciliation Pipeline Error: ${data.message}`);
    
    document.getElementById('pdf-upload-group').style.display = 'flex';
    document.getElementById('pipeline-progress').style.display = 'none';
    updateSystemStatus('online', 'Pipeline Status: Ready');
  }
}

function writeLog(message, type) {
  console.log(`[${type}] ${message}`);
}

// --- STEP 2: PARSED PDF DATA PREVIEW ---
function renderParsedPdfTable() {
  const table = document.getElementById('pdf-parsed-table');
  table.innerHTML = '';

  if (appState.pdfHeaders.length === 0 || appState.pdfRows.length === 0) {
    table.innerHTML = '<tbody><tr><td style="text-align:center; padding:30px;">No PDF records parsed.</td></tr></tbody>';
    return;
  }

  // Create Header
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  
  const thFile = document.createElement('th');
  thFile.innerText = 'PDF Filename';
  headerRow.appendChild(thFile);

  appState.pdfHeaders.forEach(h => {
    const th = document.createElement('th');
    th.innerText = h;
    headerRow.appendChild(th);
  });

  // Action header
  const thAction = document.createElement('th');
  thAction.innerText = 'Action';
  thAction.style.textAlign = 'center';
  headerRow.appendChild(thAction);

  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Create Body
  const tbody = document.createElement('tbody');
  appState.pdfRows.forEach(row => {
    const tr = document.createElement('tr');
    
    const tdFile = document.createElement('td');
    tdFile.innerText = row.filename;
    tdFile.style.fontWeight = '600';
    tr.appendChild(tdFile);

    appState.pdfHeaders.forEach(h => {
      const td = document.createElement('td');
      td.innerText = row[h] !== undefined ? row[h] : '';
      tr.appendChild(td);
    });

    // Action cell with "View Text" button
    const tdAction = document.createElement('td');
    tdAction.style.textAlign = 'center';
    
    const viewBtn = document.createElement('button');
    viewBtn.className = 'success-btn';
    viewBtn.style.padding = '4px 10px';
    viewBtn.style.fontSize = '12px';
    viewBtn.style.minHeight = 'auto';
    viewBtn.style.display = 'inline-flex';
    viewBtn.style.alignItems = 'center';
    viewBtn.style.gap = '4px';
    viewBtn.innerHTML = `<i data-lucide="eye" style="width: 12px; height: 12px;"></i> View Text`;
    
    viewBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      showParsedTextModal(row.filename, row.rawMarkdown, row);
    });
    
    tdAction.appendChild(viewBtn);
    tr.appendChild(tdAction);

    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  lucide.createIcons();
}

function showParsedTextModal(filename, rawMarkdown, rowData) {
  const modal = document.getElementById('parsed-text-modal');
  const title = document.getElementById('modal-title');
  const pdfContainer = document.getElementById('modal-pdf-viewer-container');
  const renderedContent = document.getElementById('modal-tab-rendered-content');
  const rawContent = document.getElementById('modal-tab-raw-content');
  const jsonContent = document.getElementById('modal-tab-json-content');
  
  if (!modal) return;

  // Set title
  if (title) {
    title.innerHTML = `<i data-lucide="file-text" style="color: var(--cyan); margin-right: 8px;"></i> ${filename} - Document Layout`;
  }

  // Clear previous blob url if exists
  if (currentPdfBlobUrl) {
    URL.revokeObjectURL(currentPdfBlobUrl);
    currentPdfBlobUrl = null;
  }

  // Load PDF into Left Pane
  if (pdfContainer) {
    pdfContainer.innerHTML = '';
    // Find matching local File object from appState.pdfFiles
    const file = appState.pdfFiles.find(f => f.name === filename);
    if (file) {
      currentPdfBlobUrl = URL.createObjectURL(file);
      // Create iframe
      const iframe = document.createElement('iframe');
      iframe.src = currentPdfBlobUrl;
      iframe.style.width = '100%';
      iframe.style.height = '100%';
      iframe.style.border = 'none';
      pdfContainer.appendChild(iframe);
    } else {
      pdfContainer.innerHTML = `
        <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; height:100%; color:var(--text-muted); gap:10px; padding:20px; text-align:center;">
          <i data-lucide="alert-circle" style="width:40px; height:40px; color:var(--text-muted);"></i>
          <span>Source PDF file not found in current browser cache.</span>
          <span style="font-size:12px;">Ensure the file is uploaded in Step 1.</span>
        </div>
      `;
    }
  }

  // Load Rendered Markdown
  if (renderedContent) {
    renderedContent.innerHTML = parseMarkdown(rawMarkdown);
  }

  // Load Raw Markdown Source
  if (rawContent) {
    rawContent.innerText = rawMarkdown;
  }

  // Load Extracted JSON (excluding rawMarkdown and filename for clean view)
  if (jsonContent && rowData) {
    const cleanJson = { ...rowData };
    delete cleanJson.rawMarkdown;
    delete cleanJson.filename;
    jsonContent.innerText = JSON.stringify(cleanJson, null, 2);
  }

  // Reset tab active state to "Rendered Layout" on open
  const tabs = document.querySelectorAll('.pane-tab');
  tabs.forEach(tab => {
    if (tab.dataset.tab === 'rendered') {
      tab.classList.add('active');
      tab.style.color = '#fff';
      tab.style.borderBottomColor = 'var(--cyan)';
      tab.style.fontWeight = '600';
    } else {
      tab.classList.remove('active');
      tab.style.color = 'var(--text-muted)';
      tab.style.borderBottomColor = 'transparent';
      tab.style.fontWeight = '500';
    }
  });

  const contentPanels = document.querySelectorAll('.tab-content-panel');
  contentPanels.forEach(panel => {
    if (panel.id === 'modal-tab-rendered-content') {
      panel.style.display = 'block';
    } else {
      panel.style.display = 'none';
    }
  });

  // Display modal
  modal.style.display = 'flex';
  lucide.createIcons();
}

// --- STEP 3: EXCEL UPLOAD AND MAPPING HANDLERS ---
function handleExcelFileSelect(file) {
  if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
    alert("Please upload a valid Excel spreadsheet (.xlsx or .xls)");
    return;
  }

  appState.excelFile = file;

  document.getElementById('excel-file-name').innerText = file.name;
  document.getElementById('excel-file-size').innerText = formatBytes(file.size);
  document.getElementById('excel-drop-zone').style.display = 'none';
  document.getElementById('excel-file-banner').style.display = 'flex';

  uploadExcelToServer(file);
}

async function uploadExcelToServer(file) {
  const formData = new FormData();
  formData.append('excelFile', file);

  try {
    document.getElementById('excel-preview-area').style.display = 'block';
    const previewTable = document.getElementById('excel-preview-table');
    previewTable.innerHTML = `<tr><td style="text-align:center; padding: 30px;">
      <div class="loading-spinner"></div>
      <div style="margin-top:10px;">Reading spreadsheet structure...</div>
    </td></tr>`;

    const res = await fetch('/api/upload-excel', {
      method: 'POST',
      body: formData
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Failed to read spreadsheet');
    }

    const data = await res.json();
    appState.excelHeaders = data.headers;
    appState.excelPreview = data.preview;

    renderExcelPreviewTable();
    populateMappingLayout();

    document.getElementById('mapping-container').style.display = 'block';
    validateReconciliationTrigger();

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

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  appState.excelHeaders.forEach(h => {
    const th = document.createElement('th');
    th.innerText = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

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

function resetExcelUpload() {
  appState.excelFile = null;
  appState.excelHeaders = [];
  appState.excelPreview = [];
  appState.mappings = {};
  appState.matchKey = '';

  document.getElementById('excel-drop-zone').style.display = 'flex';
  document.getElementById('excel-file-banner').style.display = 'none';
  document.getElementById('excel-preview-area').style.display = 'none';
  document.getElementById('mapping-container').style.display = 'none';
  document.getElementById('btn-run-reconciliation').disabled = true;
  document.getElementById('excel-input').value = '';
}

function populateMappingLayout() {
  const mappingGrid = document.getElementById('mapping-grid');
  const matchKeySelect = document.getElementById('match-key-select');

  mappingGrid.innerHTML = '';
  matchKeySelect.innerHTML = '<option value="" disabled selected>Select unique matching key...</option>';

  appState.mappings = {};

  appState.pdfHeaders.forEach(pdfHeader => {
    appState.mappings[pdfHeader] = '';

    // Create Mapping Row Container
    const row = document.createElement('div');
    row.className = 'mapping-row';
    row.style.display = 'flex';
    row.style.alignItems = 'center';
    row.style.gap = '15px';
    row.style.marginBottom = '12px';

    const pdfCol = document.createElement('div');
    pdfCol.style.flex = '1';
    pdfCol.style.fontWeight = '500';
    pdfCol.innerHTML = `<span class="pdf-field-tag">${pdfHeader}</span>`;

    const arrowCol = document.createElement('div');
    arrowCol.innerHTML = '<i data-lucide="arrow-right" style="color:var(--cyan)"></i>';

    const excelCol = document.createElement('div');
    excelCol.style.flex = '1.2';
    excelCol.className = 'select-wrapper';

    const select = document.createElement('select');
    select.id = `map-select-${pdfHeader.replace(/\s+/g, '-')}`;
    select.innerHTML = '<option value="" selected>Do not compare (ignore)</option>';
    
    appState.excelHeaders.forEach(eh => {
      const option = document.createElement('option');
      option.value = eh;
      option.innerText = eh;
      select.appendChild(option);
    });

    select.addEventListener('change', (e) => {
      appState.mappings[pdfHeader] = e.target.value;
      updateMatchKeyDropdownOptions();
      validateReconciliationTrigger();
    });

    excelCol.appendChild(select);
    excelCol.innerHTML += '<i data-lucide="chevron-down" class="select-chevron"></i>';
    // Reattach select event listener after innerHTML replacement
    excelCol.querySelector('select').addEventListener('change', (e) => {
      appState.mappings[pdfHeader] = e.target.value;
      updateMatchKeyDropdownOptions();
      validateReconciliationTrigger();
    });

    row.appendChild(pdfCol);
    row.appendChild(arrowCol);
    row.appendChild(excelCol);
    mappingGrid.appendChild(row);
  });

  lucide.createIcons();
}

function updateMatchKeyDropdownOptions() {
  const matchKeySelect = document.getElementById('match-key-select');
  const previousVal = appState.matchKey;

  matchKeySelect.innerHTML = '<option value="" disabled>Select unique matching key...</option>';

  let hasSelectedPrevious = false;
  
  Object.keys(appState.mappings).forEach(pdfHeader => {
    const excelHeader = appState.mappings[pdfHeader];
    // Only allow mapped fields to serve as unique match keys
    if (excelHeader) {
      const option = document.createElement('option');
      option.value = pdfHeader;
      option.innerText = `${pdfHeader} (↔ ${excelHeader})`;
      if (pdfHeader === previousVal) {
        option.selected = true;
        hasSelectedPrevious = true;
      }
      matchKeySelect.appendChild(option);
    }
  });

  if (!hasSelectedPrevious) {
    appState.matchKey = '';
    matchKeySelect.value = '';
  }
}

function validateReconciliationTrigger() {
  const btn = document.getElementById('btn-run-reconciliation');
  const matchKeySelected = appState.matchKey;
  
  // Must have selected a unique match key, and that key must have a mapped Excel header
  const isValid = matchKeySelected && appState.mappings[matchKeySelected];
  btn.disabled = !isValid;
}

// AI Auto column mapping
async function runAiAutoMapping() {
  const btn = document.getElementById('btn-ai-auto-map');
  const originalHtml = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = `<div class="loading-spinner" style="width:14px; height:14px; margin-right:6px; border-width:2px;"></div> Mapping columns...`;

  try {
    const res = await fetch('/api/auto-map-columns', {
      method: 'POST',
      headers: {
        'content-type': 'application/json'
      },
      body: JSON.stringify({
        pdfHeaders: appState.pdfHeaders,
        excelHeaders: appState.excelHeaders
      })
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'AI mapping request failed.');
    }

    const data = await res.json();
    const mapping = data.mapping;

    // Apply mappings to dropdowns and state
    Object.keys(mapping).forEach(pdfHeader => {
      const excelHeader = mapping[pdfHeader];
      if (excelHeader && appState.excelHeaders.includes(excelHeader)) {
        appState.mappings[pdfHeader] = excelHeader;
        const select = document.getElementById(`map-select-${pdfHeader.replace(/\s+/g, '-')}`);
        if (select) {
          select.value = excelHeader;
        }
      }
    });

    // Auto select first mapped column as match key
    updateMatchKeyDropdownOptions();
    const matchKeyKeys = Object.keys(appState.mappings).filter(k => appState.mappings[k]);
    if (matchKeyKeys.length > 0) {
      // Try to find a field containing 'id', 'num', or 'key'
      let keyToSelect = matchKeyKeys.find(k => k.toLowerCase().includes('id') || k.toLowerCase().includes('num') || k.toLowerCase().includes('key'));
      if (!keyToSelect) keyToSelect = matchKeyKeys[0];
      
      appState.matchKey = keyToSelect;
      document.getElementById('match-key-select').value = keyToSelect;
    }

    validateReconciliationTrigger();

  } catch (error) {
    alert("AI Auto-mapping failed: " + error.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalHtml;
  }
}

// Run Reconciliation Task
async function runReconciliation() {
  const btn = document.getElementById('btn-run-reconciliation');
  const originalHtml = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = `<div class="loading-spinner" style="width:16px; height:16px; margin-right:8px; border-width:2px;"></div> Reconciling datasets...`;

  try {
    // Block going back
    document.getElementById('step-nav-1').style.pointerEvents = 'none';
    document.getElementById('step-nav-2').style.pointerEvents = 'none';
    document.getElementById('step-nav-3').style.pointerEvents = 'none';

    // Show loading spinner in Step 4 Executive Box
    document.getElementById('executive-report-content').innerHTML = `
      <div style="display: flex; align-items: center; gap: 10px; color: #a0aec0; padding: 10px 0;">
        <div class="loading-spinner" style="width: 20px; height: 20px; border-width: 2px;"></div>
        <span>Generating auditor executive report via Claude...</span>
      </div>
    `;

    const res = await fetch('/api/reconcile-tables', {
      method: 'POST',
      headers: {
        'content-type': 'application/json'
      },
      body: JSON.stringify({
        mappings: appState.mappings,
        matchKey: appState.matchKey
      })
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Reconciliation request failed.');
    }

    const data = await res.json();
    appState.comparisonReport = data.report;
    appState.executiveSummary = data.executiveSummary;

    // Render step 4 details
    renderStep4Dashboard();
    goToStep(4);

    // Re-enable navigation clicks
    document.getElementById('step-nav-1').style.pointerEvents = 'auto';
    document.getElementById('step-nav-2').style.pointerEvents = 'auto';
    document.getElementById('step-nav-3').style.pointerEvents = 'auto';

  } catch (error) {
    alert("Reconciliation Error: " + error.message);
    document.getElementById('step-nav-1').style.pointerEvents = 'auto';
    document.getElementById('step-nav-2').style.pointerEvents = 'auto';
    document.getElementById('step-nav-3').style.pointerEvents = 'auto';
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalHtml;
  }
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

  // 3. Render Claude Executive Summary
  const mdBox = document.getElementById('executive-report-content');
  mdBox.style.background = 'rgba(255, 255, 255, 0.03)';
  mdBox.style.border = '1px solid rgba(255, 255, 255, 0.08)';
  mdBox.style.padding = '25px';
  mdBox.innerHTML = parseMarkdown(appState.executiveSummary);

  // 4. Render Detail Log Table
  renderReportTable();

  // 5. Render Unmatched Excel Table
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
    <th>Key Value (${appState.matchKey})</th>
  `;
  thead.appendChild(headerTr);
  table.appendChild(thead);

  // Create Body Rows
  const tbody = document.createElement('tbody');
  
  report.records.forEach((record, index) => {
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

    const keyVal = record.pdfMatchValue || 'N/A';

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
    const selectedHeaders = Object.keys(appState.mappings);
    selectedHeaders.forEach(header => {
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
            <span class="field-name-label">${header} ${field.excelHeader ? `(↔ ${field.excelHeader})` : ''}</span>
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
          <div class="expanded-title">Field Comparison Breakdown</div>
          <div class="details-grid">
            ${fieldsHtml}
          </div>
        </div>
      </td>
    `;

    tbody.appendChild(expTr);

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
    table.innerHTML = '<tbody><tr><td style="text-align:center; padding: 15px;">All Excel rows were matched with PDF records!</td></tr></tbody>';
    return;
  }

  // Create headers
  const thead = document.createElement('thead');
  const headerTr = document.createElement('tr');
  headerTr.innerHTML = `<th>Row Number</th>`;
  
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
                         (record.pdfMatchValue && String(record.pdfMatchValue).toLowerCase().includes(query));
    
    let matchesStatus = true;
    if (filter !== 'ALL') {
      matchesStatus = record.status === filter;
    }

    if (matchesSearch && matchesStatus) {
      row.style.display = 'table-row';
      expRow.style.display = 'none';
      arrow.classList.remove('open');
    } else {
      row.style.display = 'none';
      expRow.style.display = 'none';
      arrow.classList.remove('open');
    }
  });
}

// Simple Markdown-to-HTML parser for Claude's Auditor Report
function parseMarkdown(md) {
  if (!md) return '';
  let html = md;
  // Clean wrappers
  html = html.replace(/```markdown\s*/g, '').replace(/```\s*/g, '');
  // Blockquotes
  html = html.replace(/^> (.*)$/gim, '<blockquote style="border-left: 4px solid var(--cyan); background: rgba(56,189,248,0.05); padding: 15px 20px; margin: 15px 0; border-radius: 4px; font-style: italic; color:#e2e8f0;">$1</blockquote>');
  // Headers
  html = html.replace(/^### (.*)$/gim, '<h3 style="margin-top:25px; margin-bottom:12px; font-weight:600; color:var(--cyan); font-family:\'Montserrat\', sans-serif;">$1</h3>');
  html = html.replace(/^## (.*)$/gim, '<h2 style="margin-top:30px; margin-bottom:15px; font-weight:600; color:var(--cyan); font-family:\'Montserrat\', sans-serif; border-bottom:1px solid rgba(255,255,255,0.08); padding-bottom:8px;">$1</h2>');
  html = html.replace(/^# (.*)$/gim, '<h1 style="margin-top:35px; margin-bottom:20px; font-weight:700; color:#fff; font-family:\'Montserrat\', sans-serif;">$1</h1>');
  // Bold / Italics
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong style="color:#fff; font-weight:600;">$1</strong>');
  html = html.replace(/\*(.*?)\*/g, '<em style="color:#cbd5e1;">$1</em>');
  // Inline code
  html = html.replace(/`([^`]+)`/g, '<code style="background: rgba(255,255,255,0.1); padding: 2px 6px; border-radius: 4px; font-family: monospace; color:var(--cyan); font-size:14px;">$1</code>');
  // Bullet items
  html = html.replace(/^\s*-\s+(.*)$/gim, '<li style="margin-bottom:8px; margin-left: 20px; list-style-type: disc; color:#e2e8f0;">$1</li>');
  // Wrap list tags
  html = html.replace(/(<li style=".*?">.*?<\/li>)(?!\s*<li)/gs, '<ul style="margin: 15px 0;">$1</ul>');

  // Paragraph tags
  html = html.split('\n\n').map(p => {
    const trimmed = p.trim();
    if (trimmed.startsWith('<h') || trimmed.startsWith('<ul') || trimmed.startsWith('<block') || trimmed.startsWith('<li')) {
      return trimmed;
    }
    return `<p style="margin-bottom:15px; color:#cbd5e1;">${trimmed.replace(/\n/g, '<br>')}</p>`;
  }).join('\n');

  return html;
}

// Restart Application
function restartApplication() {
  if (confirm("Are you sure you want to start a new reconciliation project? Current results will be reset.")) {
    appState = {
      currentStep: 1,
      pdfFiles: [],
      pdfHeaders: [],
      pdfRows: [],
      excelFile: null,
      excelHeaders: [],
      excelPreview: [],
      mappings: {},
      matchKey: '',
      comparisonReport: null,
      executiveSummary: ''
    };

    resetPdfParsing();
    resetExcelUpload();
    
    document.getElementById('step-nav-1').style.pointerEvents = 'auto';
    document.getElementById('step-nav-2').style.pointerEvents = 'auto';
    document.getElementById('step-nav-3').style.pointerEvents = 'auto';

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
