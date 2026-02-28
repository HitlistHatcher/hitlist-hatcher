/* ============================================================
   HITLIST HATCHER 3.0 — APPLICATION LOGIC
   app.js

   Phase build status:
   [✓] Phase 0  — Scaffold, file upload wiring, drag-and-drop
   [ ] Phase 1  — MRRS data parser
   [ ] Phase 2  — Core readiness logic
   [ ] Phase 3  — Report renderer
   [ ] Phase 4  — Settings panel
   [ ] Phase 5  — Settings persistence (localStorage + JSON)
   [ ] Phase 6  — PDF export
   [ ] Phase 7  — Polish & testing
   [ ] Phase 8  — Deployment & distribution

   Contents:
   1.  Constants
   2.  State
   3.  DOM References
   4.  Initialization
   5.  File Upload Handling
   6.  MRRS Parser          ← built in Phase 1
   7.  Readiness Logic      ← built in Phase 2
   8.  Report Renderer      ← built in Phase 3
   9.  Settings Panel       ← built in Phase 4
   10. Settings Persistence ← built in Phase 5
   11. PDF Export           ← built in Phase 6
   12. Utilities
============================================================ */


/* ── 1. CONSTANTS ──────────────────────────────────────────── */

const APP_VERSION = '3.0.0';

// MRRS file structure — header is always row index 2 (0-based), data starts row index 3
const MRRS_HEADER_ROW = 2;
const MRRS_DATA_START_ROW = 3;
const MRRS_SHEET_NAME = 'IMR Detail';

// Column names exactly as they appear in the MRRS export header row.
// These map to the JavaScript property names used throughout the app.
// Built out fully in Phase 1.
const MRRS_COLUMNS = {
  NAME:              'Name',
  RANK:              'Rank/Rate',
  COMP_DEPT:         'Comp/Dept',
  PLATOON:           'Platoon',
  OFF_ENL:           'Off Enl Indicator',
  SEX:               'Sex',
  IMR_STATUS:        'IMR Status',
  PHA_DUE:           'PHA Due',
  PHA_DT:            'PHA Dt',
  DENTAL_EXAM_DUE:   'Dental Exam Due',
  DENTAL_EXAM_DT:    'Dental Exam Dt',
  HIV_DUE:           'HIV Test Due',
  AUDIO_DUE:         'Audio 2216 Due',
  PDHA_DUE:          'PDHA Due',
  PDHRA_DUE:         'PDHRA Due',
  TST_DUE:           'TST Due Dt',
  TST_QUEST_DUE:     'TST Quest Due',
  MAMMOGRAM_DUE:     'Mammogram Due',
  PAP_DUE:           'Pap Smear Due',
  // Immunizations — each has Req, Due, and Deferred fields
  INFLUENZA_REQ:     'Influenza Req',
  INFLUENZA_DUE:     'Influenza Due',
  INFLUENZA_DEF:     'Influenza Deferred',
  TDAP_REQ:          'Tet/Dipth Req',
  TDAP_DUE:          'Tet/Dipth Due',
  TDAP_DEF:          'Tet/Dipth Deferred',
  TYPHOID_REQ:       'Typhoid Req',
  TYPHOID_DUE:       'Typhoid Due',
  TYPHOID_DEF:       'Typhoid Deferred',
  VARICELLA_REQ:     'Varicella Req',
  VARICELLA_DUE:     'Varicella Due',
  VARICELLA_DEF:     'Varicella Deferred',
  MMR_REQ:           'MMR Req',
  MMR_DUE:           'MMR Due',
  MMR_DEF:           'MMR Deferred',
  HEPA_REQ:          'HepA Req',
  HEPA_DUE:          'HepA Due',
  HEPA_DEF:          'HepA Deferred',
  HEPB_REQ:          'HepB Req',
  HEPB_DUE:          'HepB Due',
  HEPB_DEF:          'HepB Deferred',
  TWINRIX_REQ:       'TwinRix Req',
  TWINRIX_DUE:       'TwinRix Due',
  TWINRIX_DEF:       'TwinRix Deferred',
  RABIES_REQ:        'Rabies Req',
  RABIES_DUE:        'Rabies Due',
  RABIES_DEF:        'Rabies Deferred',
  RABIES_TITER_DUE:  'Rabies Titer Due',
  CHOLERA_REQ:       'Cholera Req',
  CHOLERA_DUE:       'Cholera Due',
  CHOLERA_DEF:       'Cholera Deferred',
  JEV_REQ:           'JEV Req',
  JEV_DUE:           'JEV Due',
  JEV_DEF:           'JEV Deferred',
  MGC_REQ:           'MGC Req',
  MGC_DUE:           'MGC Due',
  MGC_DEF:           'MGC Deferred',
  POLIO_REQ:         'Polio Req',
  POLIO_DUE:         'Polio Due',
  POLIO_DEF:         'Polio Deferred',
  YF_REQ:            'Yellow Fever Req',
  YF_DUE:            'Yellow Fever Due',
  YF_DEF:            'Yellow Fever Deferred',
  ANTHRAX_REQ:       'Anthrax Req',
  ANTHRAX_DUE:       'Anthrax Due',
  ANTHRAX_DEF:       'Anthrax Deferred',
  SMALLPOX_REQ:      'Smallpox Req',
  SMALLPOX_DUE:      'Smallpox Due',
  SMALLPOX_DEF:      'Smallpox Deferred',
  ADENOVIRUS_REQ:    'Adenovirus Req',
  ADENOVIRUS_DUE:    'Adenovirus Due',
  ADENOVIRUS_DEF:    'Adenovirus Deferred',
  PNEUMO_REQ:        'Pneumococcal Req',
  PNEUMO_DUE:        'Pneumococcal Due',
  PNEUMO_DEF:        'Pneumococcal Deferred',
  HPV_REQ:           'HPV Req',
  HPV_DUE:           'HPV Due',
  HPV_DEF:           'HPV Deferred',
};

// Status codes returned by the readiness evaluator
const STATUS = {
  OVERDUE:   'OVERDUE',    // red
  DUE_SOON:  'DUE_SOON',  // yellow
  UPCOMING:  'UPCOMING',  // green
  OK:        'OK',         // not due — row may still be included if other items are due
  NA:        'NA',         // not applicable / not required
};

// localStorage key for all persisted settings
const STORAGE_KEY = 'hitlistHatcher_settings_v3';

// Column count thresholds for the live column counter warning
const COL_WARN_YELLOW = 12;
const COL_WARN_RED    = 14;

// Debounce delay for text inputs (ms)
const DEBOUNCE_MS = 500;


/* ── 2. STATE ──────────────────────────────────────────────── */

// Holds the parsed MRRS data and current app configuration.
// Populated progressively as phases are built.
const state = {
  rawData:        null,    // Raw SheetJS workbook object
  personnel:      [],      // Array of parsed personnel objects (Phase 1)
  settings:       {},      // Current user settings (Phase 4/5)
  reportGenerated: false,  // Whether a report is currently displayed
};


/* ── 3. DOM REFERENCES ─────────────────────────────────────── */

const dom = {
  uploadArea:          document.getElementById('uploadArea'),
  fileInput:           document.getElementById('fileInput'),
  uploadStatus:        document.getElementById('uploadStatus'),
  generateBtn:         document.getElementById('generateBtn'),
  columnCount:         document.getElementById('columnCount'),
  columnCounter:       document.getElementById('columnCounter'),
  exportSettingsBtn:   document.getElementById('exportSettingsBtn'),
  importSettingsInput: document.getElementById('importSettingsInput'),
  previewPlaceholder:  document.getElementById('previewPlaceholder'),
  reportOutput:        document.getElementById('reportOutput'),
  printOutput:         document.getElementById('printOutput'),
  exportBar:           document.getElementById('exportBar'),
  exportPdfBtn:        document.getElementById('exportPdfBtn'),
};


/* ── 4. INITIALIZATION ─────────────────────────────────────── */

document.addEventListener('DOMContentLoaded', () => {
  console.log(`Hitlist Hatcher ${APP_VERSION} — initializing`);

  wireUploadHandlers();
  wireExportPdfHandler();
  wireSettingsHandlers();

  // Phase 5: loadSettings() will be called here once persistence is built
  // Phase 4: initSettingsPanel() will be called here once settings UI is built

  console.log('Initialization complete. Ready for file upload.');
});


/* ── 5. FILE UPLOAD HANDLING ───────────────────────────────── */

function wireUploadHandlers() {
  // Click on upload area opens the file picker
  dom.uploadArea.addEventListener('click', () => dom.fileInput.click());

  // File selected via file picker
  dom.fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
  });

  // Drag-and-drop support
  dom.uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dom.uploadArea.classList.add('drag-over');
  });

  dom.uploadArea.addEventListener('dragleave', () => {
    dom.uploadArea.classList.remove('drag-over');
  });

  dom.uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    dom.uploadArea.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
}

function handleFile(file) {
  console.log(`File selected: ${file.name} (${formatBytes(file.size)})`);

  // Basic extension check
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    setUploadStatus('error', '✗ Please upload an Excel file (.xlsx or .xls)');
    dom.uploadArea.classList.add('upload-error');
    return;
  }

  setUploadStatus('', 'Reading file...');

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });

      // Verify this looks like an MRRS file
      if (!workbook.SheetNames.includes(MRRS_SHEET_NAME)) {
        setUploadStatus('error', `✗ Sheet "${MRRS_SHEET_NAME}" not found. Is this an MRRS IMR Detail export?`);
        dom.uploadArea.classList.add('upload-error');
        dom.uploadArea.classList.remove('upload-success');
        return;
      }

      // Store the raw workbook
      state.rawData = workbook;

      dom.uploadArea.classList.add('upload-success');
      dom.uploadArea.classList.remove('upload-error');
      setUploadStatus('success', `✓ ${file.name} loaded successfully`);

      // Enable the generate button
      dom.generateBtn.disabled = false;

      console.log('Workbook loaded. Sheets:', workbook.SheetNames);
      console.log('Phase 1 (parser) will process this data when Generate is clicked.');

      // Phase 1: parseMRRS(workbook) will be called here

    } catch (err) {
      console.error('File read error:', err);
      setUploadStatus('error', '✗ Could not read file. Please try again.');
      dom.uploadArea.classList.add('upload-error');
    }
  };

  reader.onerror = () => {
    setUploadStatus('error', '✗ File read failed. Please try again.');
  };

  reader.readAsArrayBuffer(file);
}

function setUploadStatus(type, message) {
  dom.uploadStatus.textContent = message;
  dom.uploadStatus.className = 'upload-status' + (type ? ` ${type}` : '');
}


/* ── 6. MRRS PARSER ────────────────────────────────────────── */
/*
   Built in Phase 1.
   Will parse state.rawData (SheetJS workbook) into state.personnel[].
   Each personnel object will have named properties matching MRRS_COLUMNS.
*/

// Placeholder — will be replaced in Phase 1
function parseMRRS(workbook) {
  console.log('parseMRRS() — Phase 1 not yet built.');
  return [];
}


/* ── 7. READINESS LOGIC ────────────────────────────────────── */
/*
   Built in Phase 2.
   evaluateItem(dueDate, projectionDate, thresholds) → STATUS constant
   evaluateImmunizations(person, selectedVaccines, projectionDate, thresholds) → { status, count, names[] }
*/

// Placeholder — will be replaced in Phase 2
function evaluateItem(dueDate, projectionDate, thresholds) {
  console.log('evaluateItem() — Phase 2 not yet built.');
  return STATUS.NA;
}


/* ── 8. REPORT RENDERER ────────────────────────────────────── */
/*
   Built in Phase 3.
   renderReport(personnel, settings) → injects HTML into dom.reportOutput
*/

// Placeholder — will be replaced in Phase 3
function renderReport(personnel, settings) {
  console.log('renderReport() — Phase 3 not yet built.');
}


/* ── 9. SETTINGS PANEL ─────────────────────────────────────── */
/*
   Built in Phase 4.
   All settings UI is wired here.
*/

// Placeholder — will be built in Phase 4
function initSettingsPanel() {
  console.log('initSettingsPanel() — Phase 4 not yet built.');
}

// Generate button handler — wires parser + logic + renderer together
dom.generateBtn.addEventListener('click', () => {
  if (!state.rawData) {
    alert('Please upload an MRRS file first.');
    return;
  }

  console.log('Generate clicked — Phase 1-3 will run here once built.');

  // Phases 1-3 will chain here:
  // state.personnel = parseMRRS(state.rawData);
  // const filtered  = applyFilters(state.personnel, state.settings);
  // renderReport(filtered, state.settings);
  // dom.exportBar.hidden = false;
  // dom.previewPlaceholder.hidden = true;

  // Temporary: show a placeholder message in the preview
  dom.previewPlaceholder.hidden = true;
  dom.reportOutput.innerHTML = `
    <div style="padding:40px; text-align:center; color:#9aa3b0;">
      <p style="font-size:16px; font-weight:600; margin-bottom:8px;">
        ✓ File loaded — parser not yet built
      </p>
      <p style="font-size:13px;">
        The MRRS data is in memory and ready. The report renderer will be built in Phases 1–3.
      </p>
    </div>
  `;
});


/* ── 10. SETTINGS PERSISTENCE ──────────────────────────────── */
/*
   Built in Phase 5.
   saveSettings(), loadSettings(), exportSettings(), importSettings()
*/

function wireSettingsHandlers() {
  // Export settings
  dom.exportSettingsBtn.addEventListener('click', () => {
    console.log('exportSettings() — Phase 5 not yet built.');
    alert('Settings export will be available in Phase 5.');
  });

  // Import settings
  dom.importSettingsInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      console.log('importSettings() — Phase 5 not yet built.');
      alert('Settings import will be available in Phase 5.');
    }
  });
}

// Placeholder functions — replaced in Phase 5
function saveSettings()   { /* Phase 5 */ }
function loadSettings()   { /* Phase 5 */ }
function exportSettings() { /* Phase 5 */ }
function importSettings() { /* Phase 5 */ }


/* ── 11. PDF EXPORT ────────────────────────────────────────── */

function wireExportPdfHandler() {
  dom.exportPdfBtn.addEventListener('click', () => {
    // Copy the report into the print-only area, then trigger print
    dom.printOutput.innerHTML = dom.reportOutput.innerHTML;
    window.print();
  });
}


/* ── 12. UTILITIES ─────────────────────────────────────────── */

/**
 * Formats a byte count into a human-readable string (e.g., "2.4 MB").
 */
function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Formats a JavaScript Date as abbreviated military date string (e.g., "15 JAN").
 */
function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${date.getDate()} ${months[date.getMonth()]}`;
}

/**
 * Returns the number of days between two dates (positive = future, negative = past).
 */
function daysDiff(dateA, dateB) {
  const msPerDay = 1000 * 60 * 60 * 24;
  return Math.round((dateA - dateB) / msPerDay);
}

/**
 * Creates a debounced version of a function that delays invocation
 * until after `delay` ms have elapsed since the last call.
 */
function debounce(fn, delay) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

/**
 * Updates the live column counter element based on current count.
 * Called whenever the user checks/unchecks report items.
 */
function updateColumnCounter(count) {
  dom.columnCount.textContent = count;
  dom.columnCounter.classList.remove('warn-yellow', 'warn-red');

  if (count >= COL_WARN_RED) {
    dom.columnCounter.classList.add('warn-red');
    dom.columnCounter.title = 'Wide reports may be difficult to read when printed.';
  } else if (count >= COL_WARN_YELLOW) {
    dom.columnCounter.classList.add('warn-yellow');
    dom.columnCounter.title = 'Approaching maximum readable column count for landscape printing.';
  } else {
    dom.columnCounter.title = '';
  }
}
