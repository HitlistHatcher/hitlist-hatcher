/* ============================================================
   HITLIST HATCHER 3.0 — APPLICATION LOGIC
   app.js

   Phase build status:
   [✓] Phase 0  — Scaffold, file upload wiring, drag-and-drop
   [✓] Phase 1  — MRRS data parser
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
   6.  MRRS Parser          [✓ Phase 1]
   7.  Readiness Logic      ← built in Phase 2
   8.  Report Renderer      ← built in Phase 3
   9.  Settings Panel       ← built in Phase 4
   10. Settings Persistence ← built in Phase 5
   11. PDF Export           ← built in Phase 6
   12. Utilities
============================================================ */


/* ── 1. CONSTANTS ──────────────────────────────────────────── */

const APP_VERSION = '3.0.0';

// MRRS file structure
const MRRS_HEADER_ROW  = 2;   // 0-based: row index 2 = spreadsheet row 3
const MRRS_DATA_START  = 3;   // 0-based: row index 3 = spreadsheet row 4
const MRRS_SHEET_NAME  = 'IMR Detail';

// Exact column header strings as they appear in the MRRS export.
// If MRRS ever adds/renames a column, only this object needs updating.
const COL = {
  NAME:             'Name',
  RANK:             'Rank/Rate',
  COMP_DEPT:        'Comp/Dept',
  PLATOON:          'Platoon',
  OFF_ENL:          'Off Enl Indicator',
  SEX:              'Sex',
  IMR_STATUS:       'IMR Status',
  DEPLOYING:        'Deploying',

  // Core readiness
  PHA_DT:           'PHA Dt',
  PHA_DUE:          'PHA Due',
  DENTAL_DT:        'Dental Exam Dt',
  DENTAL_DUE:       'Dental Exam Due',
  HIV_DT:           'HIV Test Dt',
  HIV_DUE:          'HIV Test Due',
  AUDIO_DT:         'Audio 2216 Dt',
  AUDIO_DUE:        'Audio 2216 Due',

  // Deployment health
  PDHA_DUE:         'PDHA Due',
  PDHRA_DUE:        'PDHRA Due',

  // TB / TST
  TST_DUE:          'TST Due Dt',
  TST_QUEST_DUE:    'TST Quest Due',

  // Women's health (displayed as "Well-Woman" for privacy)
  MAMMOGRAM_DUE:    'Mammogram Due',
  PAP_DUE:          'Pap Smear Due',

  // Immunizations — triplets of (Req, Due, Deferred)
  INFLUENZA_REQ:    'Influenza Req',
  INFLUENZA_DUE:    'Influenza Due',
  INFLUENZA_DEF:    'Influenza Deferred',

  TDAP_REQ:         'Tet/Dipth Req',
  TDAP_DUE:         'Tet/Dipth Due',
  TDAP_DEF:         'Tet/Dipth Deferred',

  TYPHOID_REQ:      'Typhoid Req',
  TYPHOID_DUE:      'Typhoid Due',
  TYPHOID_DEF:      'Typhoid Deferred',

  VARICELLA_REQ:    'Varicella Req',
  VARICELLA_DUE:    'Varicella Due',
  VARICELLA_DEF:    'Varicella Deferred',

  MMR_REQ:          'MMR Req',
  MMR_DUE:          'MMR Due',
  MMR_DEF:          'MMR Deferred',

  HEPA_REQ:         'HepA Req',
  HEPA_DUE:         'HepA Due',
  HEPA_DEF:         'HepA Deferred',

  HEPB_REQ:         'HepB Req',
  HEPB_DUE:         'HepB Due',
  HEPB_DEF:         'HepB Deferred',

  TWINRIX_REQ:      'TwinRix Req',
  TWINRIX_DUE:      'TwinRix Due',
  TWINRIX_DEF:      'TwinRix Deferred',

  RABIES_REQ:       'Rabies Req',
  RABIES_DUE:       'Rabies Due',
  RABIES_DEF:       'Rabies Deferred',
  RABIES_TITER_DUE: 'Rabies Titer Due',

  CHOLERA_REQ:      'Cholera Req',
  CHOLERA_DUE:      'Cholera Due',
  CHOLERA_DEF:      'Cholera Deferred',

  JEV_REQ:          'JEV Req',
  JEV_DUE:          'JEV Due',
  JEV_DEF:          'JEV Deferred',

  MGC_REQ:          'MGC Req',
  MGC_DUE:          'MGC Due',
  MGC_DEF:          'MGC Deferred',

  POLIO_REQ:        'Polio Req',
  POLIO_DUE:        'Polio Due',
  POLIO_DEF:        'Polio Deferred',

  YF_REQ:           'Yellow Fever Req',
  YF_DUE:           'Yellow Fever Due',
  YF_DEF:           'Yellow Fever Deferred',

  ANTHRAX_REQ:      'Anthrax Req',
  ANTHRAX_DUE:      'Anthrax Due',
  ANTHRAX_DEF:      'Anthrax Deferred',

  SMALLPOX_REQ:     'Smallpox Req',
  SMALLPOX_DUE:     'Smallpox Due',
  SMALLPOX_DEF:     'Smallpox Deferred',

  ADENOVIRUS_REQ:   'Adenovirus Req',
  ADENOVIRUS_DUE:   'Adenovirus Due',
  ADENOVIRUS_DEF:   'Adenovirus Deferred',

  PNEUMO_REQ:       'Pneumococcal Req',
  PNEUMO_DUE:       'Pneumococcal Due',
  PNEUMO_DEF:       'Pneumococcal Deferred',

  HPV_REQ:          'HPV Req',
  HPV_DUE:          'HPV Due',
  HPV_DEF:          'HPV Deferred',
};

// Status codes returned by the readiness evaluator (Phase 2)
const STATUS = {
  OVERDUE:  'OVERDUE',   // red
  DUE_SOON: 'DUE_SOON', // yellow
  UPCOMING: 'UPCOMING', // green
  OK:       'OK',        // not due within any warning window
  NA:       'NA',        // not applicable / not required
};

// localStorage key
const STORAGE_KEY = 'hitlistHatcher_settings_v3';

// Column counter warning thresholds
const COL_WARN_YELLOW = 12;
const COL_WARN_RED    = 14;

// Debounce delay for text inputs (ms)
const DEBOUNCE_MS = 500;


/* ── 2. STATE ──────────────────────────────────────────────── */

const state = {
  rawData:         null,   // SheetJS workbook object
  personnel:       [],     // Parsed personnel array (populated by parseMRRS)
  colIndex:        {},     // Maps column header string → array index (built by parseMRRS)
  parseStats:      {},     // Summary statistics from the last parse
  settings:        {},     // Current user settings (Phase 4/5)
  reportGenerated: false,
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
  console.log('Ready for file upload.');
});


/* ── 5. FILE UPLOAD HANDLING ───────────────────────────────── */

function wireUploadHandlers() {
  dom.uploadArea.addEventListener('click', () => dom.fileInput.click());

  dom.fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
  });

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

  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    setUploadStatus('error', '✗ Please upload an Excel file (.xlsx or .xls)');
    dom.uploadArea.classList.add('upload-error');
    return;
  }

  setUploadStatus('', 'Reading file...');

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const data     = new Uint8Array(e.target.result);
      // cellDates:true  → SheetJS returns JS Date objects for date cells
      // raw:false       → Formatted string values for non-date cells
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });

      if (!workbook.SheetNames.includes(MRRS_SHEET_NAME)) {
        setUploadStatus('error', `✗ Sheet "${MRRS_SHEET_NAME}" not found. Is this an MRRS IMR Detail export?`);
        dom.uploadArea.classList.add('upload-error');
        dom.uploadArea.classList.remove('upload-success');
        return;
      }

      state.rawData = workbook;

      // ── PHASE 1: Parse immediately on load ──────────────────
      const parsed = parseMRRS(workbook);

      if (parsed.personnel.length === 0) {
        setUploadStatus('error', '✗ No personnel rows found. Please check this is an MRRS IMR Detail export.');
        dom.uploadArea.classList.add('upload-error');
        return;
      }

      state.personnel  = parsed.personnel;
      state.colIndex   = parsed.colIndex;
      state.parseStats = parsed.stats;

      dom.uploadArea.classList.add('upload-success');
      dom.uploadArea.classList.remove('upload-error');
      setUploadStatus('success',
        `✓ ${file.name} — ${parsed.personnel.length} personnel loaded ` +
        `(${parsed.stats.officers} officers, ${parsed.stats.enlisted} enlisted)`
      );

      dom.generateBtn.disabled = false;

      // Log parse results to console for verification
      logParseResults(parsed);

    } catch (err) {
      console.error('File read error:', err);
      setUploadStatus('error', '✗ Could not read file. Please try again.');
      dom.uploadArea.classList.add('upload-error');
    }
  };

  reader.onerror = () => setUploadStatus('error', '✗ File read failed. Please try again.');
  reader.readAsArrayBuffer(file);
}

function setUploadStatus(type, message) {
  dom.uploadStatus.textContent  = message;
  dom.uploadStatus.className    = 'upload-status' + (type ? ` ${type}` : '');
}


/* ── 6. MRRS PARSER ────────────────────────────────────────── */
/*
   parseMRRS(workbook)
   ───────────────────
   Reads the 'IMR Detail' sheet from a SheetJS workbook object.
   Returns { personnel[], colIndex{}, stats{} }.

   personnel[] — one object per row, with properties named after
                 the COL constant keys (e.g. person.NAME, person.PHA_DUE)
   colIndex{}  — maps each header string to its column array index
                 (kept in state for diagnostics and future flexibility)
   stats{}     — summary counts for the upload status message
*/

function parseMRRS(workbook) {
  const sheet = workbook.Sheets[MRRS_SHEET_NAME];

  // Convert the sheet to a 2D array.
  // header:1 means SheetJS uses the first row as keys — we do NOT want that
  // because our header is row 3, not row 1.
  // So we use sheet_to_json with header:1 to get a raw 2D array, then
  // manually handle the header row ourselves.
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,          // Return array of arrays (not objects)
    defval: '',         // Empty cells become empty string, not undefined
    raw: false,         // Use formatted values for strings
  });

  // ── 6a. Build the column index map ─────────────────────────
  // Row index 2 (spreadsheet row 3) is the header row.
  const headerRow = rows[MRRS_HEADER_ROW] || [];
  const colIndex  = {};

  headerRow.forEach((headerVal, i) => {
    if (headerVal !== null && headerVal !== undefined && headerVal !== '') {
      colIndex[String(headerVal).trim()] = i;
    }
  });

  // Warn in console if any expected columns are missing
  const missingCols = [];
  Object.values(COL).forEach(colName => {
    if (!(colName in colIndex)) missingCols.push(colName);
  });
  if (missingCols.length > 0) {
    console.warn(`parseMRRS: ${missingCols.length} expected column(s) not found in this file:`, missingCols);
  }

  // ── 6b. Helper: get a cell value by COL key ─────────────────
  // Returns the raw value (string, Date, number, or '') for a
  // given COL constant key and row array.
  function getCell(row, colKey) {
    const idx = colIndex[COL[colKey]];
    if (idx === undefined) return '';
    const val = row[idx];
    return val === null || val === undefined ? '' : val;
  }

  // ── 6c. Helper: parse a date cell ───────────────────────────
  // SheetJS with cellDates:true returns JS Date objects for date
  // cells when it can detect them. However, some date cells arrive
  // as formatted strings (e.g. "10/30/2018") or Excel serial numbers.
  // This function normalises all three cases to a JS Date or null.
  function parseDate(val) {
    if (!val || val === '') return null;

    // Already a Date object (SheetJS cellDates:true)
    if (val instanceof Date) {
      return isNaN(val.getTime()) ? null : val;
    }

    // Numeric — Excel serial date (days since 1900-01-00)
    if (typeof val === 'number') {
      // SheetJS helper converts serial to Date
      const d = XLSX.SSF.parse_date_code(val);
      if (d) return new Date(d.y, d.m - 1, d.d);
      return null;
    }

    // String — attempt to parse common formats
    if (typeof val === 'string') {
      const trimmed = val.trim();
      if (trimmed === '') return null;

      // ISO format: "2018-10-30T00:00:00.000Z" or "2018-10-30"
      const iso = Date.parse(trimmed);
      if (!isNaN(iso)) return new Date(iso);

      // MM/DD/YYYY
      const mdy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (mdy) return new Date(parseInt(mdy[3]), parseInt(mdy[1]) - 1, parseInt(mdy[2]));

      // DD-MON-YYYY (e.g. "30-OCT-2018")
      const dmy = trimmed.match(/^(\d{1,2})-([A-Z]{3})-(\d{4})$/i);
      if (dmy) {
        const months = { JAN:0,FEB:1,MAR:2,APR:3,MAY:4,JUN:5,JUL:6,AUG:7,SEP:8,OCT:9,NOV:10,DEC:11 };
        const m = months[dmy[2].toUpperCase()];
        if (m !== undefined) return new Date(parseInt(dmy[3]), m, parseInt(dmy[1]));
      }
    }

    return null;
  }

  // ── 6d. Helper: dental due date calculation ─────────────────
  // Two methods per confirmed project decision:
  //   MRRS method  — use the MRRS-supplied Dental Exam Due date
  //   12-Month     — add 365 days to the Dental Exam Dt date
  // The settings toggle (Phase 4) will set useMrrsDate.
  // For now we default to 12-Month (365-day) as per the project plan.
  function dentalDueDate(row, useMrrsDate = false) {
    if (useMrrsDate) {
      return parseDate(getCell(row, 'DENTAL_DUE'));
    } else {
      const examDt = parseDate(getCell(row, 'DENTAL_DT'));
      if (!examDt) return null;
      const due = new Date(examDt);
      due.setDate(due.getDate() + 365);
      return due;
    }
  }

  // ── 6e. Helper: normalise boolean-ish string fields ─────────
  // MRRS uses values like "Yes", "No", "Y", "N", blank, etc.
  // Returns true if the value suggests "required" or "yes".
  function isYes(val) {
    if (!val && val !== 0) return false;
    const s = String(val).trim().toLowerCase();
    return s === 'yes' || s === 'y' || s === '1' || s === 'true';
  }

  // Returns true if the value suggests "deferred".
  function isDeferred(val) {
    if (!val && val !== 0) return false;
    const s = String(val).trim().toLowerCase();
    return s === 'yes' || s === 'y' || s === '1' || s === 'true' || s === 'deferred';
  }

  // ── 6f. Helper: section label ────────────────────────────────
  // Combines Comp/Dept and Platoon into a single readable label.
  function sectionLabel(row) {
    const dept     = String(getCell(row, 'COMP_DEPT')).trim();
    const platoon  = String(getCell(row, 'PLATOON')).trim();
    if (dept && platoon) return `${dept}-${platoon}`;
    return dept || platoon || '';
  }

  // ── 6g. Parse all data rows ──────────────────────────────────
  const personnel = [];
  let officers = 0;
  let enlisted = 0;
  let skipped  = 0;

  for (let i = MRRS_DATA_START; i < rows.length; i++) {
    const row = rows[i];

    // Skip completely empty rows (can appear at end of sheet)
    const name = String(getCell(row, 'NAME')).trim();
    if (!name || name === '') {
      skipped++;
      continue;
    }

    const offEnl = String(getCell(row, 'OFF_ENL')).trim();
    const rank   = String(getCell(row, 'RANK')).trim();

    // Chief Warrant Officers may have a blank Off Enl Indicator in some
    // MRRS exports. If the indicator is blank, fall back to checking
    // whether Rank/Rate contains "CWO". CWOs are treated as officers:
    // they appear in Combined and Officer-only reports, not Enlisted-only.
    const isOfficer = offEnl.toLowerCase().includes('officer') ||
                  (offEnl === '' && rank.toUpperCase().includes('CWO'));

    if (isOfficer) officers++; else enlisted++;

    // Build the personnel object.
    // All date fields are normalised to JS Date objects or null.
    // All string fields are trimmed.
    const person = {
      // Identity
      name:        name,
      rank:        String(getCell(row, 'RANK')).trim(),
      section:     sectionLabel(row),
      offEnl:      isOfficer ? 'Officer' : 'Enlisted',
      sex:         String(getCell(row, 'SEX')).trim(),
      imrStatus:   String(getCell(row, 'IMR_STATUS')).trim(),
      deploying:   isYes(getCell(row, 'DEPLOYING')),

      // Core readiness — due dates
      phaDue:      parseDate(getCell(row, 'PHA_DUE')),
      dentalDue:   dentalDueDate(row, false),   // false = 12-month method
      hivDue:      parseDate(getCell(row, 'HIV_DUE')),
      audioDue:    parseDate(getCell(row, 'AUDIO_DUE')),

      // Deployment
      pdhaDue:     parseDate(getCell(row, 'PDHA_DUE')),
      pdhraDue:    parseDate(getCell(row, 'PDHRA_DUE')),

      // TB / TST
      tstDue:      parseDate(getCell(row, 'TST_DUE')),
      tstQuestDue: parseDate(getCell(row, 'TST_QUEST_DUE')),

      // Women's health
      mammogramDue: parseDate(getCell(row, 'MAMMOGRAM_DUE')),
      papDue:        parseDate(getCell(row, 'PAP_DUE')),

      // Immunizations — each stored as { req, due, deferred }
      // The readiness evaluator (Phase 2) will use this structure.
      immunizations: {
        influenza:  { req: isYes(getCell(row,'INFLUENZA_REQ')),  due: parseDate(getCell(row,'INFLUENZA_DUE')),  deferred: isDeferred(getCell(row,'INFLUENZA_DEF')),  label: 'Influenza' },
        tdap:       { req: isYes(getCell(row,'TDAP_REQ')),       due: parseDate(getCell(row,'TDAP_DUE')),       deferred: isDeferred(getCell(row,'TDAP_DEF')),       label: 'TDap' },
        typhoid:    { req: isYes(getCell(row,'TYPHOID_REQ')),    due: parseDate(getCell(row,'TYPHOID_DUE')),    deferred: isDeferred(getCell(row,'TYPHOID_DEF')),    label: 'Typhoid' },
        varicella:  { req: isYes(getCell(row,'VARICELLA_REQ')), due: parseDate(getCell(row,'VARICELLA_DUE')), deferred: isDeferred(getCell(row,'VARICELLA_DEF')), label: 'Varicella' },
        mmr:        { req: isYes(getCell(row,'MMR_REQ')),        due: parseDate(getCell(row,'MMR_DUE')),        deferred: isDeferred(getCell(row,'MMR_DEF')),        label: 'MMR' },
        hepa:       { req: isYes(getCell(row,'HEPA_REQ')),       due: parseDate(getCell(row,'HEPA_DUE')),       deferred: isDeferred(getCell(row,'HEPA_DEF')),       label: 'Hep A' },
        hepb:       { req: isYes(getCell(row,'HEPB_REQ')),       due: parseDate(getCell(row,'HEPB_DUE')),       deferred: isDeferred(getCell(row,'HEPB_DEF')),       label: 'Hep B' },
        twinrix:    { req: isYes(getCell(row,'TWINRIX_REQ')),    due: parseDate(getCell(row,'TWINRIX_DUE')),    deferred: isDeferred(getCell(row,'TWINRIX_DEF')),    label: 'TwinRix' },
        rabies:     { req: isYes(getCell(row,'RABIES_REQ')),     due: parseDate(getCell(row,'RABIES_DUE')),     deferred: isDeferred(getCell(row,'RABIES_DEF')),     label: 'Rabies' },
        rabiesTiter:{ req: isYes(getCell(row,'RABIES_REQ')),     due: parseDate(getCell(row,'RABIES_TITER_DUE')), deferred: false,                                   label: 'Rabies Titer' },
        cholera:    { req: isYes(getCell(row,'CHOLERA_REQ')),    due: parseDate(getCell(row,'CHOLERA_DUE')),    deferred: isDeferred(getCell(row,'CHOLERA_DEF')),    label: 'Cholera' },
        jev:        { req: isYes(getCell(row,'JEV_REQ')),        due: parseDate(getCell(row,'JEV_DUE')),        deferred: isDeferred(getCell(row,'JEV_DEF')),        label: 'JEV' },
        mgc:        { req: isYes(getCell(row,'MGC_REQ')),        due: parseDate(getCell(row,'MGC_DUE')),        deferred: isDeferred(getCell(row,'MGC_DEF')),        label: 'MGC' },
        polio:      { req: isYes(getCell(row,'POLIO_REQ')),      due: parseDate(getCell(row,'POLIO_DUE')),      deferred: isDeferred(getCell(row,'POLIO_DEF')),      label: 'Polio' },
        yellowFever:{ req: isYes(getCell(row,'YF_REQ')),         due: parseDate(getCell(row,'YF_DUE')),         deferred: isDeferred(getCell(row,'YF_DEF')),         label: 'Yellow Fever' },
        anthrax:    { req: isYes(getCell(row,'ANTHRAX_REQ')),    due: parseDate(getCell(row,'ANTHRAX_DUE')),    deferred: isDeferred(getCell(row,'ANTHRAX_DEF')),    label: 'Anthrax' },
        smallpox:   { req: isYes(getCell(row,'SMALLPOX_REQ')),   due: parseDate(getCell(row,'SMALLPOX_DUE')),   deferred: isDeferred(getCell(row,'SMALLPOX_DEF')),   label: 'Smallpox' },
        adenovirus: { req: isYes(getCell(row,'ADENOVIRUS_REQ')), due: parseDate(getCell(row,'ADENOVIRUS_DUE')), deferred: isDeferred(getCell(row,'ADENOVIRUS_DEF')), label: 'Adenovirus' },
        pneumo:     { req: isYes(getCell(row,'PNEUMO_REQ')),     due: parseDate(getCell(row,'PNEUMO_DUE')),     deferred: isDeferred(getCell(row,'PNEUMO_DEF')),     label: 'Pneumococcal' },
        hpv:        { req: isYes(getCell(row,'HPV_REQ')),        due: parseDate(getCell(row,'HPV_DUE')),        deferred: isDeferred(getCell(row,'HPV_DEF')),        label: 'HPV' },
      },
    };

    personnel.push(person);
  }

  const stats = {
    total:    personnel.length,
    officers,
    enlisted,
    skipped,
    columns:  Object.keys(colIndex).length,
  };

  return { personnel, colIndex, stats };
}


/* ── 6h. CONSOLE VERIFICATION LOGGER ──────────────────────── */
/*
   Logs a structured summary to the browser console after parsing
   so you can verify the data looks correct against your MRRS file.
   Open DevTools (F12) → Console after uploading a file.
   This function is safe to leave in — it only runs in the browser
   and produces no visible UI output.
*/

function logParseResults(parsed) {
  const { personnel, stats } = parsed;

  console.group(`%cHitlist Hatcher — Parse Results`, 'color:#003087; font-weight:bold; font-size:14px;');

  console.log(`%cSummary`, 'font-weight:bold;');
  console.table({
    'Total personnel': stats.total,
    'Officers':        stats.officers,
    'Enlisted':        stats.enlisted,
    'Rows skipped':    stats.skipped,
    'Columns found':   stats.columns,
  });

  console.log(`%cFirst 5 personnel (identity fields)`, 'font-weight:bold;');
  console.table(
    personnel.slice(0, 5).map(p => ({
      name:      p.name,
      rank:      p.rank,
      section:   p.section,
      category:  p.offEnl,
      imrStatus: p.imrStatus,
    }))
  );

  console.log(`%cFirst 5 personnel (core readiness due dates — full dates for verification)`, 'font-weight:bold;');
  console.table(
    personnel.slice(0, 5).map(p => ({
      name:      p.name,
      phaDue:    p.phaDue    ? formatDateFull(p.phaDue)    : 'null',
      dentalDue: p.dentalDue ? formatDateFull(p.dentalDue) : 'null',
      hivDue:    p.hivDue    ? formatDateFull(p.hivDue)    : 'null',
      audioDue:  p.audioDue  ? formatDateFull(p.audioDue)  : 'null',
    }))
  );

  console.log(`%cFirst 5 personnel (immunization sample — Influenza & TDap)`, 'font-weight:bold;');
  console.table(
    personnel.slice(0, 5).map(p => ({
      name:           p.name,
      influenza_req:  p.immunizations.influenza.req,
      influenza_due:  p.immunizations.influenza.due ? formatDate(p.immunizations.influenza.due) : 'null',
      influenza_def:  p.immunizations.influenza.deferred,
      tdap_req:       p.immunizations.tdap.req,
      tdap_due:       p.immunizations.tdap.due ? formatDate(p.immunizations.tdap.due) : 'null',
    }))
  );

  console.log(
    `%cFull personnel array available as: window._hhPersonnel`,
    'color:#1e8449; font-style:italic;'
  );
  window._hhPersonnel = personnel; // Available in console for ad-hoc inspection

  console.groupEnd();
}


/* ── 7. READINESS LOGIC ────────────────────────────────────── */
/*  Built in Phase 2  */

function evaluateItem(dueDate, projectionDate, thresholds) {
  console.log('evaluateItem() — Phase 2 not yet built.');
  return STATUS.NA;
}


/* ── 8. REPORT RENDERER ────────────────────────────────────── */
/*  Built in Phase 3  */

function renderReport(personnel, settings) {
  console.log('renderReport() — Phase 3 not yet built.');
}


/* ── 9. SETTINGS PANEL ─────────────────────────────────────── */
/*  Built in Phase 4  */

function initSettingsPanel() {
  console.log('initSettingsPanel() — Phase 4 not yet built.');
}

// Generate button — will chain Phases 1-3 together
dom.generateBtn.addEventListener('click', () => {
  if (!state.rawData || state.personnel.length === 0) {
    alert('Please upload an MRRS file first.');
    return;
  }

  console.log(`Generate clicked — ${state.personnel.length} personnel in memory. Phases 2-3 will process these once built.`);

  // Temporary preview message until Phase 3 is built
  dom.previewPlaceholder.hidden = true;
  dom.reportOutput.innerHTML = `
    <div style="
      background:#fff;
      border-radius:6px;
      padding:40px;
      text-align:center;
      box-shadow:0 2px 10px rgba(0,0,0,0.1);
    ">
      <div style="font-size:48px; margin-bottom:16px;">✓</div>
      <p style="font-size:18px; font-weight:700; color:#003087; margin-bottom:8px;">
        Phase 1 Complete — Parser Working
      </p>
      <p style="font-size:14px; color:#444; margin-bottom:16px;">
        ${state.personnel.length} personnel parsed successfully
        (${state.parseStats.officers} officers,
         ${state.parseStats.enlisted} enlisted).
      </p>
      <p style="font-size:13px; color:#9aa3b0;">
        Open the browser console (F12 → Console) to inspect the parsed data.<br />
        The report renderer will be built in Phase 3.
      </p>
    </div>
  `;
});


/* ── 10. SETTINGS PERSISTENCE ──────────────────────────────── */
/*  Built in Phase 5  */

function wireSettingsHandlers() {
  dom.exportSettingsBtn.addEventListener('click', () => {
    alert('Settings export will be available in Phase 5.');
  });
  dom.importSettingsInput.addEventListener('change', () => {
    alert('Settings import will be available in Phase 5.');
  });
}

function saveSettings()   { /* Phase 5 */ }
function loadSettings()   { /* Phase 5 */ }
function exportSettings() { /* Phase 5 */ }
function importSettings() { /* Phase 5 */ }


/* ── 11. PDF EXPORT ────────────────────────────────────────── */

function wireExportPdfHandler() {
  dom.exportPdfBtn.addEventListener('click', () => {
    dom.printOutput.innerHTML = dom.reportOutput.innerHTML;
    window.print();
  });
}


/* ── 12. UTILITIES ─────────────────────────────────────────── */

function formatBytes(bytes) {
  if (bytes < 1024)        return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${date.getDate()} ${months[date.getMonth()]}`;
}

// Full date format for console verification and diagnostics only.
// The hit list display uses formatDate() (no year) for column width efficiency.
function formatDateFull(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

function daysDiff(dateA, dateB) {
  const msPerDay = 1000 * 60 * 60 * 24;
  return Math.round((dateA - dateB) / msPerDay);
}

function debounce(fn, delay) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}

function updateColumnCounter(count) {
  dom.columnCount.textContent = count;
  dom.columnCounter.classList.remove('warn-yellow', 'warn-red');
  if (count >= COL_WARN_RED) {
    dom.columnCounter.classList.add('warn-red');
  } else if (count >= COL_WARN_YELLOW) {
    dom.columnCounter.classList.add('warn-yellow');
  }
}
