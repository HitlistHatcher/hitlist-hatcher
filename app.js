/* ============================================================
   HITLIST HATCHER 3.0 — APPLICATION LOGIC
   app.js

   Phase build status:
   [✓] Phase 0  — Scaffold, file upload wiring, drag-and-drop
   [✓] Phase 1  — MRRS data parser
   [✓] Phase 2  — Core readiness logic
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
   7.  Readiness Logic      [✓ Phase 2]
   8.  Report Renderer      ← built in Phase 3
   9.  Settings Panel       ← built in Phase 4
   10. Settings Persistence ← built in Phase 5
   11. PDF Export           ← built in Phase 6
   12. Utilities
============================================================ */


/* ── 1. CONSTANTS ──────────────────────────────────────────── */

const APP_VERSION = '3.0.0';

const MRRS_HEADER_ROW = 2;
const MRRS_DATA_START = 3;
const MRRS_SHEET_NAME = 'IMR Detail';

const COL = {
  NAME:             'Name',
  RANK:             'Rank/Rate',
  COMP_DEPT:        'Comp/Dept',
  PLATOON:          'Platoon',
  OFF_ENL:          'Off Enl Indicator',
  SEX:              'Sex',
  IMR_STATUS:       'IMR Status',
  DEPLOYING:        'Deploying',

  PHA_DT:           'PHA Dt',
  PHA_DUE:          'PHA Due',
  DENTAL_DT:        'Dental Exam Dt',
  DENTAL_DUE:       'Dental Exam Due',
  DENTAL_CLASS:     'Dental Cond Code',
  HIV_DT:           'HIV Test Dt',
  HIV_DUE:          'HIV Test Due',
  AUDIO_DT:         'Audio 2216 Dt',
  AUDIO_DUE:        'Audio 2216 Due',

  PDHA_DUE:         'PDHA Due',
  PDHRA_DUE:        'PDHRA Due',

  TST_DUE:          'TST Due Dt',
  TST_QUEST_DUE:    'TST Quest Due',

  MAMMOGRAM_DUE:    'Mammogram Due',
  PAP_DUE:          'Pap Smear Due',

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

// Status codes — order matters for severity comparison (index = severity)
const STATUS = {
  NA:       'NA',       // not applicable / not required
  OK:       'OK',       // not due within any warning window
  UPCOMING: 'UPCOMING', // green — due within green window
  DUE_SOON: 'DUE_SOON', // yellow — due within yellow window
  OVERDUE:  'OVERDUE',  // red — overdue or Class 3/4 dental
};

// Severity rank — higher number = more urgent.
// Used to find the worst status across multiple items.
const SEVERITY = {
  NA:       0,
  OK:       1,
  UPCOMING: 2,
  DUE_SOON: 3,
  OVERDUE:  4,
};

// Default warning thresholds (days). All user-configurable in Phase 4.
const DEFAULT_THRESHOLDS = {
  yellow: 7,   // due within 7 days → yellow
  green:  30,  // due within 30 days → green
  // red = anything at or past 0 days (today or overdue)
};

const STORAGE_KEY      = 'hitlistHatcher_settings_v3';
const COL_WARN_YELLOW  = 12;
const COL_WARN_RED     = 14;
const DEBOUNCE_MS      = 500;

// All immunization keys — used by the evaluator and renderer.
// Order here controls display order in individual-column mode.
const IMMUNIZATION_KEYS = [
  'influenza', 'tdap', 'typhoid', 'varicella', 'mmr',
  'hepa', 'hepb', 'twinrix', 'rabies', 'rabiesTiter',
  'cholera', 'jev', 'mgc', 'polio', 'yellowFever',
  'anthrax', 'smallpox', 'adenovirus', 'pneumo', 'hpv',
];

// Default checked state for immunizations in the settings panel (Phase 4)
const IMMUNIZATION_DEFAULTS = {
  influenza:   true,
  tdap:        true,
  typhoid:     false,
  varicella:   false,
  mmr:         false,
  hepa:        false,
  hepb:        false,
  twinrix:     false,
  rabies:      false,
  rabiesTiter: false,
  cholera:     false,
  jev:         false,
  mgc:         false,
  polio:       false,
  yellowFever: false,
  anthrax:     false,
  smallpox:    false,
  adenovirus:  false,
  pneumo:      false,
  hpv:         false,
};

// Human-readable labels for each immunization key
const IMMUNIZATION_LABELS = {
  influenza:   'Influenza',
  tdap:        'TDap',
  typhoid:     'Typhoid',
  varicella:   'Varicella',
  mmr:         'MMR',
  hepa:        'Hep A',
  hepb:        'Hep B',
  twinrix:     'TwinRix',
  rabies:      'Rabies',
  rabiesTiter: 'Rabies Titer',
  cholera:     'Cholera',
  jev:         'JEV',
  mgc:         'MGC',
  polio:       'Polio',
  yellowFever: 'Yellow Fever',
  anthrax:     'Anthrax',
  smallpox:    'Smallpox',
  adenovirus:  'Adenovirus',
  pneumo:      'Pneumococcal',
  hpv:         'HPV',
};


/* ── 2. STATE ──────────────────────────────────────────────── */

const state = {
  rawData:         null,
  personnel:       [],
  colIndex:        {},
  parseStats:      {},
  filteredResults: [],   // Output of applyFilters() — what the renderer displays
  settings:        {},
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
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });

      if (!workbook.SheetNames.includes(MRRS_SHEET_NAME)) {
        setUploadStatus('error', `✗ Sheet "${MRRS_SHEET_NAME}" not found. Is this an MRRS IMR Detail export?`);
        dom.uploadArea.classList.add('upload-error');
        dom.uploadArea.classList.remove('upload-success');
        return;
      }

      state.rawData = workbook;

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
  dom.uploadStatus.textContent = message;
  dom.uploadStatus.className   = 'upload-status' + (type ? ` ${type}` : '');
}


/* ── 6. MRRS PARSER ────────────────────────────────────────── */

function parseMRRS(workbook) {
  const sheet = workbook.Sheets[MRRS_SHEET_NAME];
  const rows  = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw:    false,
  });

  // Build column index map from header row
  const headerRow = rows[MRRS_HEADER_ROW] || [];
  const colIndex  = {};
  headerRow.forEach((val, i) => {
    if (val !== null && val !== undefined && val !== '') {
      colIndex[String(val).trim()] = i;
    }
  });

  // Warn about missing columns
  const missingCols = Object.values(COL).filter(c => !(c in colIndex));
  if (missingCols.length > 0) {
    console.warn(`parseMRRS: ${missingCols.length} expected column(s) not found:`, missingCols);
  }

  function getCell(row, colKey) {
    const idx = colIndex[COL[colKey]];
    if (idx === undefined) return '';
    const val = row[idx];
    return (val === null || val === undefined) ? '' : val;
  }

  function parseDate(val) {
    if (!val || val === '') return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
    if (typeof val === 'number') {
      const d = XLSX.SSF.parse_date_code(val);
      return d ? new Date(d.y, d.m - 1, d.d) : null;
    }
    if (typeof val === 'string') {
      const trimmed = val.trim();
      if (!trimmed) return null;
      const iso = Date.parse(trimmed);
      if (!isNaN(iso)) return new Date(iso);
      const mdy = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (mdy) return new Date(parseInt(mdy[3]), parseInt(mdy[1]) - 1, parseInt(mdy[2]));
      const dmy = trimmed.match(/^(\d{1,2})-([A-Z]{3})-(\d{4})$/i);
      if (dmy) {
        const months = { JAN:0,FEB:1,MAR:2,APR:3,MAY:4,JUN:5,JUL:6,AUG:7,SEP:8,OCT:9,NOV:10,DEC:11 };
        const m = months[dmy[2].toUpperCase()];
        if (m !== undefined) return new Date(parseInt(dmy[3]), m, parseInt(dmy[1]));
      }
    }
    return null;
  }

  function dentalDueDate(row, useMrrsDate = false) {
    if (useMrrsDate) return parseDate(getCell(row, 'DENTAL_DUE'));
    const examDt = parseDate(getCell(row, 'DENTAL_DT'));
    if (!examDt) return null;
    const due = new Date(examDt);
    due.setDate(due.getDate() + 365);
    return due;
  }

  function parseDentalClass(val) {
    if (!val || String(val).trim() === '') return 4;
    const n = parseInt(String(val).trim(), 10);
    return (n >= 1 && n <= 4) ? n : 4;
  }

  function isYes(val) {
    if (!val && val !== 0) return false;
    const s = String(val).trim().toLowerCase();
    return s === 'yes' || s === 'y' || s === '1' || s === 'true';
  }

  function isDeferred(val) {
    if (!val && val !== 0) return false;
    const s = String(val).trim().toLowerCase();
    return s === 'yes' || s === 'y' || s === '1' || s === 'true' || s === 'deferred';
  }

  function sectionLabel(row) {
    const dept    = String(getCell(row, 'COMP_DEPT')).trim();
    const platoon = String(getCell(row, 'PLATOON')).trim();
    if (dept && platoon) return `${dept}-${platoon}`;
    return dept || platoon || '';
  }

  const personnel = [];
  let officers = 0, enlisted = 0, skipped = 0;

  for (let i = MRRS_DATA_START; i < rows.length; i++) {
    const row  = rows[i];
    const name = String(getCell(row, 'NAME')).trim();
    if (!name) { skipped++; continue; }

    const offEnl    = String(getCell(row, 'OFF_ENL')).trim();
    const rank      = String(getCell(row, 'RANK')).trim();
    const isOfficer = offEnl.toLowerCase().includes('officer') ||
                      (offEnl === '' && rank.toUpperCase().includes('CWO'));

    if (isOfficer) officers++; else enlisted++;

    const person = {
      name,
      rank,
      section:     sectionLabel(row),
      offEnl:      isOfficer ? 'Officer' : 'Enlisted',
      sex:         String(getCell(row, 'SEX')).trim(),
      imrStatus:   String(getCell(row, 'IMR_STATUS')).trim(),
      deploying:   isYes(getCell(row, 'DEPLOYING')),

      phaDue:      parseDate(getCell(row, 'PHA_DUE')),
      dentalDue:   dentalDueDate(row, false),
      dentalClass: parseDentalClass(getCell(row, 'DENTAL_CLASS')),
      hivDue:      parseDate(getCell(row, 'HIV_DUE')),
      audioDue:    parseDate(getCell(row, 'AUDIO_DUE')),

      pdhaDue:     parseDate(getCell(row, 'PDHA_DUE')),
      pdhraDue:    parseDate(getCell(row, 'PDHRA_DUE')),

      tstDue:      parseDate(getCell(row, 'TST_DUE')),
      tstQuestDue: parseDate(getCell(row, 'TST_QUEST_DUE')),

      mammogramDue: parseDate(getCell(row, 'MAMMOGRAM_DUE')),
      papDue:       parseDate(getCell(row, 'PAP_DUE')),

      immunizations: {
        influenza:   { req: isYes(getCell(row,'INFLUENZA_REQ')),  due: parseDate(getCell(row,'INFLUENZA_DUE')),  deferred: isDeferred(getCell(row,'INFLUENZA_DEF')),  label: 'Influenza' },
        tdap:        { req: isYes(getCell(row,'TDAP_REQ')),       due: parseDate(getCell(row,'TDAP_DUE')),       deferred: isDeferred(getCell(row,'TDAP_DEF')),       label: 'TDap' },
        typhoid:     { req: isYes(getCell(row,'TYPHOID_REQ')),    due: parseDate(getCell(row,'TYPHOID_DUE')),    deferred: isDeferred(getCell(row,'TYPHOID_DEF')),    label: 'Typhoid' },
        varicella:   { req: isYes(getCell(row,'VARICELLA_REQ')), due: parseDate(getCell(row,'VARICELLA_DUE')), deferred: isDeferred(getCell(row,'VARICELLA_DEF')), label: 'Varicella' },
        mmr:         { req: isYes(getCell(row,'MMR_REQ')),        due: parseDate(getCell(row,'MMR_DUE')),        deferred: isDeferred(getCell(row,'MMR_DEF')),        label: 'MMR' },
        hepa:        { req: isYes(getCell(row,'HEPA_REQ')),       due: parseDate(getCell(row,'HEPA_DUE')),       deferred: isDeferred(getCell(row,'HEPA_DEF')),       label: 'Hep A' },
        hepb:        { req: isYes(getCell(row,'HEPB_REQ')),       due: parseDate(getCell(row,'HEPB_DUE')),       deferred: isDeferred(getCell(row,'HEPB_DEF')),       label: 'Hep B' },
        twinrix:     { req: isYes(getCell(row,'TWINRIX_REQ')),    due: parseDate(getCell(row,'TWINRIX_DUE')),    deferred: isDeferred(getCell(row,'TWINRIX_DEF')),    label: 'TwinRix' },
        rabies:      { req: isYes(getCell(row,'RABIES_REQ')),     due: parseDate(getCell(row,'RABIES_DUE')),     deferred: isDeferred(getCell(row,'RABIES_DEF')),     label: 'Rabies' },
        rabiesTiter: { req: isYes(getCell(row,'RABIES_REQ')),     due: parseDate(getCell(row,'RABIES_TITER_DUE')), deferred: false,                                   label: 'Rabies Titer' },
        cholera:     { req: isYes(getCell(row,'CHOLERA_REQ')),    due: parseDate(getCell(row,'CHOLERA_DUE')),    deferred: isDeferred(getCell(row,'CHOLERA_DEF')),    label: 'Cholera' },
        jev:         { req: isYes(getCell(row,'JEV_REQ')),        due: parseDate(getCell(row,'JEV_DUE')),        deferred: isDeferred(getCell(row,'JEV_DEF')),        label: 'JEV' },
        mgc:         { req: isYes(getCell(row,'MGC_REQ')),        due: parseDate(getCell(row,'MGC_DUE')),        deferred: isDeferred(getCell(row,'MGC_DEF')),        label: 'MGC' },
        polio:       { req: isYes(getCell(row,'POLIO_REQ')),      due: parseDate(getCell(row,'POLIO_DUE')),      deferred: isDeferred(getCell(row,'POLIO_DEF')),      label: 'Polio' },
        yellowFever: { req: isYes(getCell(row,'YF_REQ')),         due: parseDate(getCell(row,'YF_DUE')),         deferred: isDeferred(getCell(row,'YF_DEF')),         label: 'Yellow Fever' },
        anthrax:     { req: isYes(getCell(row,'ANTHRAX_REQ')),    due: parseDate(getCell(row,'ANTHRAX_DUE')),    deferred: isDeferred(getCell(row,'ANTHRAX_DEF')),    label: 'Anthrax' },
        smallpox:    { req: isYes(getCell(row,'SMALLPOX_REQ')),   due: parseDate(getCell(row,'SMALLPOX_DUE')),   deferred: isDeferred(getCell(row,'SMALLPOX_DEF')),   label: 'Smallpox' },
        adenovirus:  { req: isYes(getCell(row,'ADENOVIRUS_REQ')), due: parseDate(getCell(row,'ADENOVIRUS_DUE')), deferred: isDeferred(getCell(row,'ADENOVIRUS_DEF')), label: 'Adenovirus' },
        pneumo:      { req: isYes(getCell(row,'PNEUMO_REQ')),     due: parseDate(getCell(row,'PNEUMO_DUE')),     deferred: isDeferred(getCell(row,'PNEUMO_DEF')),     label: 'Pneumococcal' },
        hpv:         { req: isYes(getCell(row,'HPV_REQ')),        due: parseDate(getCell(row,'HPV_DUE')),        deferred: isDeferred(getCell(row,'HPV_DEF')),        label: 'HPV' },
      },
    };

    personnel.push(person);
  }

  return {
    personnel,
    colIndex,
    stats: { total: personnel.length, officers, enlisted, skipped, columns: Object.keys(colIndex).length },
  };
}


/* ── 6h. CONSOLE VERIFICATION LOGGER ──────────────────────── */

function logParseResults(parsed) {
  const { personnel, stats } = parsed;

  console.group('%cHitlist Hatcher — Parse Results', 'color:#003087;font-weight:bold;font-size:14px;');

  console.log('%cSummary', 'font-weight:bold;');
  console.table({
    'Total personnel': stats.total,
    'Officers':        stats.officers,
    'Enlisted':        stats.enlisted,
    'Rows skipped':    stats.skipped,
    'Columns found':   stats.columns,
  });

  console.log('%cFirst 5 personnel (identity fields)', 'font-weight:bold;');
  console.table(personnel.slice(0, 5).map(p => ({
    name: p.name, rank: p.rank, section: p.section,
    category: p.offEnl, dentalClass: p.dentalClass, imrStatus: p.imrStatus,
  })));

  console.log('%cFirst 5 personnel (core readiness due dates — full dates for verification)', 'font-weight:bold;');
  console.table(personnel.slice(0, 5).map(p => ({
    name:      p.name,
    phaDue:    p.phaDue    ? formatDateFull(p.phaDue)    : 'null',
    dentalDue: p.dentalDue ? formatDateFull(p.dentalDue) : 'null',
    hivDue:    p.hivDue    ? formatDateFull(p.hivDue)    : 'null',
    audioDue:  p.audioDue  ? formatDateFull(p.audioDue)  : 'null',
  })));

  console.log('%cFirst 5 personnel (Influenza & TDap immunizations)', 'font-weight:bold;');
  console.table(personnel.slice(0, 5).map(p => ({
    name:          p.name,
    influenza_req: p.immunizations.influenza.req,
    influenza_due: p.immunizations.influenza.due ? formatDateFull(p.immunizations.influenza.due) : 'null',
    influenza_def: p.immunizations.influenza.deferred,
    tdap_req:      p.immunizations.tdap.req,
    tdap_due:      p.immunizations.tdap.due ? formatDateFull(p.immunizations.tdap.due) : 'null',
  })));

  console.log('%cFull array: window._hhPersonnel', 'color:#1e8449;font-style:italic;');
  window._hhPersonnel = personnel;

  console.groupEnd();
}


/* ── 7. READINESS LOGIC ────────────────────────────────────── */

/*
  7a. evaluateItem(dueDate, projectionDate, thresholds)
  ─────────────────────────────────────────────────────
  The core comparator. Takes a due date (JS Date or null), a projection
  date (JS Date), and a thresholds object { yellow, green } (days).
  Returns a STATUS constant.

  Logic:
    - null due date                          → NA
    - daysUntilDue <= 0  (today or past)     → OVERDUE (red)
    - daysUntilDue <= thresholds.yellow      → DUE_SOON (yellow)
    - daysUntilDue <= thresholds.green       → UPCOMING (green)
    - daysUntilDue >  thresholds.green       → OK (not shown unless other items due)

  Note: daysDiff(dueDate, projectionDate) is positive when the due date
  is in the future and negative when it is in the past.
*/
function evaluateItem(dueDate, projectionDate, thresholds) {
  if (!dueDate || !(dueDate instanceof Date) || isNaN(dueDate)) return STATUS.NA;

  const days = daysDiff(dueDate, projectionDate);

  if (days <= 0)                   return STATUS.OVERDUE;
  if (days <= thresholds.yellow)   return STATUS.DUE_SOON;
  if (days <= thresholds.green)    return STATUS.UPCOMING;
  return STATUS.OK;
}


/*
  7b. evaluateDental(person, projectionDate, thresholds, useMrrsDate)
  ────────────────────────────────────────────────────────────────────
  Dedicated dental evaluator. Handles Class 3/4 override logic.

  Returns { status, displayText }
    - displayText is what appears in the cell.
    - For Class 3/4: displayText is "Class 3" or "Class 4", status is OVERDUE.
    - For Class 1/2: displayText is the formatted due date (set by renderer),
      status is from evaluateItem().

  useMrrsDate: if true, uses the MRRS-supplied due date directly.
               if false (default), uses the 365-day date already stored
               in person.dentalDue by the parser.
*/
function evaluateDental(person, projectionDate, thresholds, useMrrsDate = false) {
  const cls = person.dentalClass;

  if (cls === 3) return { status: STATUS.OVERDUE, displayText: 'Class 3' };
  if (cls === 4) return { status: STATUS.OVERDUE, displayText: 'Class 4' };

  // Class 1 or 2: evaluate by due date
  // The 365-day adjusted date is already stored in person.dentalDue by the parser.
  // If useMrrsDate is true the renderer will need the raw MRRS date —
  // but since we store only the 365-day date in the person object, the
  // settings toggle in Phase 4 will call parseMRRS with the correct flag.
  // For now person.dentalDue is always the 365-day date.
  const status = evaluateItem(person.dentalDue, projectionDate, thresholds);
  return { status, displayText: null }; // null = renderer shows the formatted date
}


/*
  7c. evaluateImmunization(immObj, projectionDate, thresholds)
  ─────────────────────────────────────────────────────────────
  Evaluates a single immunization object { req, due, deferred, label }.
  Returns STATUS.

  Logic:
    - not required (req === false)   → NA
    - deferred                       → NA (deferred = not currently actionable)
    - due date null                  → OVERDUE (required but no date on record)
    - otherwise                      → evaluateItem()
*/
function evaluateImmunization(immObj, projectionDate, thresholds) {
  if (!immObj.req)      return STATUS.NA;
  if (immObj.deferred)  return STATUS.NA;
  if (!immObj.due)      return STATUS.OVERDUE; // required but no date — flag it
  return evaluateItem(immObj.due, projectionDate, thresholds);
}


/*
  7d. evaluateImmunizations(person, selectedKeys, projectionDate, thresholds)
  ─────────────────────────────────────────────────────────────────────────────
  Aggregates immunization status across all selected vaccines for one person.
  Used by both grouped mode (returns summary) and individual mode (returns
  per-vaccine statuses).

  selectedKeys: array of immunization key strings that are checked in settings
                (e.g. ['influenza', 'tdap'])

  Returns {
    worstStatus,          // highest-severity STATUS across all selected vaccines
    count,                // number of selected vaccines that are OVERDUE or DUE_SOON or UPCOMING
    dueNames,             // array of label strings for vaccines that are not NA/OK (for tooltip)
    perVaccine,           // object keyed by vaccine key → { status, due } for individual mode
  }
*/
function evaluateImmunizations(person, selectedKeys, projectionDate, thresholds) {
  let worstStatus = STATUS.NA;
  let count       = 0;
  const dueNames  = [];
  const perVaccine = {};

  selectedKeys.forEach(key => {
    const immObj = person.immunizations[key];
    if (!immObj) return;

    const status = evaluateImmunization(immObj, projectionDate, thresholds);
    perVaccine[key] = { status, due: immObj.due };

    if (SEVERITY[status] > SEVERITY[worstStatus]) worstStatus = status;

    if (status === STATUS.OVERDUE || status === STATUS.DUE_SOON || status === STATUS.UPCOMING) {
      count++;
      dueNames.push(immObj.label);
    }
  });

  return { worstStatus, count, dueNames, perVaccine };
}


/*
  7e. evaluatePerson(person, settings, projectionDate)
  ─────────────────────────────────────────────────────
  Runs the full readiness evaluation for one person against the
  current settings configuration. Returns a result object that the
  renderer (Phase 3) consumes directly.

  settings (placeholder values used until Phase 4 builds the UI):
  {
    items: {
      pha, dental, hiv, audio, pdha, pdhra, tst, tstQuest,
      wellWoman, immunizations
    },
    immunizationKeys: [],     // which vaccines are selected
    thresholds: { yellow, green },
    dentalUseMrrsDate: false,
  }

  Returns {
    person,                   // reference to the original person object
    overallStatus,            // worst status across all evaluated items
    items: {
      pha:         { status, displayText },
      dental:      { status, displayText },
      hiv:         { status, displayText },
      audio:       { status, displayText },
      pdha:        { status, displayText },
      pdhra:       { status, displayText },
      tst:         { status, displayText },
      tstQuest:    { status, displayText },
      wellWoman:   { status, displayText },
      immunizations: {
        grouped:   { status, count, dueNames },   // for grouped display mode
        perVaccine: { [key]: { status, due } },   // for individual column mode
      }
    },
    isDue,   // boolean — true if this person should appear on the hit list
  }
*/
function evaluatePerson(person, settings, projectionDate) {
  const t = settings.thresholds || DEFAULT_THRESHOLDS;

  // ── Evaluate each selected item ────────────────────────────

  // PHA
  const phaResult = settings.items.pha
    ? { status: evaluateItem(person.phaDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  // Dental (has its own evaluator due to Class 3/4 logic)
  const dentalResult = settings.items.dental
    ? evaluateDental(person, projectionDate, t, settings.dentalUseMrrsDate || false)
    : { status: STATUS.NA, displayText: null };

  // HIV
  const hivResult = settings.items.hiv
    ? { status: evaluateItem(person.hivDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  // Audiogram
  const audioResult = settings.items.audio
    ? { status: evaluateItem(person.audioDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  // PDHA
  const pdhaResult = settings.items.pdha
    ? { status: evaluateItem(person.pdhaDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  // PDHRA
  const pdhraResult = settings.items.pdhra
    ? { status: evaluateItem(person.pdhraDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  // TST
  const tstResult = settings.items.tst
    ? { status: evaluateItem(person.tstDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  // TST Quest
  const tstQuestResult = settings.items.tstQuest
    ? { status: evaluateItem(person.tstQuestDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  // Well-Woman (Mammogram + Pap Smear — worst of the two, women only)
  let wellWomanResult = { status: STATUS.NA, displayText: null };
  if (settings.items.wellWoman && person.sex.toLowerCase() === 'female') {
    const mamStatus = evaluateItem(person.mammogramDue, projectionDate, t);
    const papStatus = evaluateItem(person.papDue, projectionDate, t);
    const worst = SEVERITY[mamStatus] >= SEVERITY[papStatus] ? mamStatus : papStatus;
    wellWomanResult = { status: worst, displayText: null };
  }

  // Immunizations
  const immKeys   = settings.immunizationKeys || [];
  const immResult = evaluateImmunizations(person, immKeys, projectionDate, t);

  const immunizationsResult = {
    grouped:    { status: immResult.worstStatus, count: immResult.count, dueNames: immResult.dueNames },
    perVaccine: immResult.perVaccine,
  };

  // ── Determine overall (worst) status ───────────────────────
  const allStatuses = [
    phaResult.status,
    dentalResult.status,
    hivResult.status,
    audioResult.status,
    pdhaResult.status,
    pdhraResult.status,
    tstResult.status,
    tstQuestResult.status,
    wellWomanResult.status,
    immResult.worstStatus,
  ];

  const overallStatus = allStatuses.reduce((worst, s) =>
    SEVERITY[s] > SEVERITY[worst] ? s : worst,
    STATUS.NA
  );

  // Person appears on the hit list if any item is OVERDUE, DUE_SOON, or UPCOMING
  const isDue = SEVERITY[overallStatus] >= SEVERITY[STATUS.UPCOMING];

  return {
    person,
    overallStatus,
    items: {
      pha:           phaResult,
      dental:        dentalResult,
      hiv:           hivResult,
      audio:         audioResult,
      pdha:          pdhaResult,
      pdhra:         pdhraResult,
      tst:           tstResult,
      tstQuest:      tstQuestResult,
      wellWoman:     wellWomanResult,
      immunizations: immunizationsResult,
    },
    isDue,
  };
}


/*
  7f. applyFilters(personnel, settings, projectionDate)
  ──────────────────────────────────────────────────────
  1. Evaluates every person via evaluatePerson().
  2. Filters to only those where isDue === true.
  3. Applies officer/enlisted filter.
  4. Sorts by name or by section-then-name.
  Returns { results[], stats{} } where results[] is the array the
  renderer will iterate over, and stats{} feeds the readiness summary.
*/
function applyFilters(personnel, settings, projectionDate) {
  // Evaluate everyone (stats need the full roster, not just the filtered list)
  const allEvaluated = personnel.map(p => evaluatePerson(p, settings, projectionDate));

  // Compute summary stats from full roster IMR status field
  const stats = computeReadinessStats(personnel);

  // Filter by officer/enlisted preference
  const offEnlFilter = settings.offEnlFilter || 'combined';
  let filtered = allEvaluated.filter(r => {
    if (!r.isDue) return false;
    if (offEnlFilter === 'combined')  return true;
    if (offEnlFilter === 'officer')   return r.person.offEnl === 'Officer';
    if (offEnlFilter === 'enlisted')  return r.person.offEnl === 'Enlisted';
    return true;
  });

  // Sort
  const sortBy = settings.sortBy || 'name';
  filtered.sort((a, b) => {
    if (sortBy === 'section') {
      const secCmp = a.person.section.localeCompare(b.person.section);
      if (secCmp !== 0) return secCmp;
    }
    return a.person.name.localeCompare(b.person.name);
  });

  return { results: filtered, stats };
}


/*
  7g. computeReadinessStats(personnel)
  ──────────────────────────────────────
  Counts IMR status values from the full personnel roster.
  These numbers feed the readiness summary in the report header.

  MRRS IMR Status values observed in the data:
    "Fully Medically Ready"
    "Not Medically Ready"
    "Partially Medically Ready"
    "Indeterminate"

  Returns { total, fullyReady, notReady, partial, indeterminate,
            fullyReadyPct, notReadyPct, partialPct, indeterminatePct }
*/
function computeReadinessStats(personnel) {
  let fullyReady    = 0;
  let notReady      = 0;
  let partial       = 0;
  let indeterminate = 0;

  personnel.forEach(p => {
    const s = (p.imrStatus || '').toLowerCase();
    if (s.includes('fully'))          fullyReady++;
    else if (s.includes('not'))       notReady++;
    else if (s.includes('partial'))   partial++;
    else                               indeterminate++;
  });

  const total = personnel.length;
  const pct   = n => total > 0 ? ((n / total) * 100).toFixed(1) : '0.0';

  return {
    total,
    fullyReady,
    notReady,
    partial,
    indeterminate,
    fullyReadyPct:    pct(fullyReady),
    notReadyPct:      pct(notReady),
    partialPct:       pct(partial),
    indeterminatePct: pct(indeterminate),
  };
}


/* ── 8. REPORT RENDERER ────────────────────────────────────── */
/*  Built in Phase 3  */

function renderReport(results, stats, settings) {
  console.log('renderReport() — Phase 3 not yet built.');
}


/* ── 9. SETTINGS PANEL ─────────────────────────────────────── */
/*  Built in Phase 4  */

function initSettingsPanel() {
  console.log('initSettingsPanel() — Phase 4 not yet built.');
}


/*
  Generate button — now chains Phases 1 and 2 together.
  Uses temporary hardcoded settings until Phase 4 builds the UI.
  Logs the Phase 2 output to the console for verification.
*/
dom.generateBtn.addEventListener('click', () => {
  if (!state.rawData || state.personnel.length === 0) {
    alert('Please upload an MRRS file first.');
    return;
  }

  // ── Temporary settings (replaced by UI in Phase 4) ─────────
  const tempSettings = {
    items: {
      pha:       true,
      dental:    true,
      hiv:       true,
      audio:     true,
      pdha:      false,
      pdhra:     false,
      tst:       false,
      tstQuest:  false,
      wellWoman: false,
      immunizations: true,
    },
    immunizationKeys:  ['influenza', 'tdap'],
    thresholds:        { yellow: 7, green: 30 },
    offEnlFilter:      'combined',
    sortBy:            'name',
    immunDisplayMode:  'grouped',
    dentalUseMrrsDate: false,
  };

  const projectionDate = new Date(); // today — replaced by date picker in Phase 4

  // Run Phase 2
  const { results, stats } = applyFilters(state.personnel, tempSettings, projectionDate);
  state.filteredResults = results;

  // Log Phase 2 output to console for verification
  logPhase2Results(results, stats, projectionDate);

  // Temporary preview card
  dom.previewPlaceholder.hidden = true;
  dom.reportOutput.innerHTML = `
    <div style="
      background:#fff; border-radius:6px; padding:40px;
      text-align:center; box-shadow:0 2px 10px rgba(0,0,0,0.1);
    ">
      <div style="font-size:48px; margin-bottom:16px;">✓</div>
      <p style="font-size:18px; font-weight:700; color:#003087; margin-bottom:8px;">
        Phase 2 Complete — Readiness Logic Working
      </p>
      <p style="font-size:14px; color:#444; margin-bottom:8px;">
        ${results.length} personnel flagged as due or upcoming
        out of ${state.personnel.length} total.
      </p>
      <p style="font-size:13px; color:#444; margin-bottom:16px;">
        Force Readiness: ${stats.fullyReadyPct}% Fully Ready &nbsp;|&nbsp;
        ${stats.notReadyPct}% Not Ready &nbsp;|&nbsp;
        ${stats.partialPct}% Partial &nbsp;|&nbsp;
        ${stats.indeterminatePct}% Indeterminate
      </p>
      <p style="font-size:13px; color:#9aa3b0;">
        Open the browser console (F12 → Console) to inspect evaluation results.<br />
        The report renderer will be built in Phase 3.
      </p>
    </div>
  `;
});


/* ── Phase 2 console verification logger ───────────────────── */

function logPhase2Results(results, stats, projectionDate) {
  console.group('%cHitlist Hatcher — Phase 2 Results', 'color:#003087;font-weight:bold;font-size:14px;');

  console.log(`%cProjection date: ${formatDateFull(projectionDate)}`, 'font-style:italic;');

  console.log('%cReadiness Summary (full roster)', 'font-weight:bold;');
  console.table({
    'Total':          stats.total,
    'Fully Ready':    `${stats.fullyReady} (${stats.fullyReadyPct}%)`,
    'Not Ready':      `${stats.notReady} (${stats.notReadyPct}%)`,
    'Partial':        `${stats.partial} (${stats.partialPct}%)`,
    'Indeterminate':  `${stats.indeterminate} (${stats.indeterminatePct}%)`,
  });

  console.log(`%c${results.length} personnel flagged (due or upcoming)`, 'font-weight:bold;');
  console.table(results.slice(0, 10).map(r => ({
    name:          r.person.name,
    rank:          r.person.rank,
    overallStatus: r.overallStatus,
    pha:           r.items.pha.status,
    dental:        r.items.dental.displayText || r.items.dental.status,
    hiv:           r.items.hiv.status,
    audio:         r.items.audio.status,
    immGrouped:    `${r.items.immunizations.grouped.count} due (${r.items.immunizations.grouped.status})`,
    immTooltip:    r.items.immunizations.grouped.dueNames.join(', ') || '—',
  })));

  console.log('%cFull results array: window._hhResults', 'color:#1e8449;font-style:italic;');
  window._hhResults = results;
  window._hhStats   = stats;

  console.groupEnd();
}


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
  if (bytes < 1024)         return `${bytes} B`;
  if (bytes < 1024 * 1024)  return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

// Abbreviated date for hit list cells (e.g. "15 JAN")
function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${date.getDate()} ${months[date.getMonth()]}`;
}

// Full date for console verification only (e.g. "15 JAN 2026")
function formatDateFull(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

// Returns days between two dates. Positive = dateA is in the future relative to dateB.
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
  if (count >= COL_WARN_RED)         dom.columnCounter.classList.add('warn-red');
  else if (count >= COL_WARN_YELLOW) dom.columnCounter.classList.add('warn-yellow');
}
