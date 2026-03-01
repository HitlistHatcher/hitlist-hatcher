/* ============================================================
   HITLIST HATCHER 3.0 — APPLICATION LOGIC
   app.js

   Phase build status:
   [✓] Phase 0  — Scaffold, file upload wiring, drag-and-drop
   [✓] Phase 1  — MRRS data parser
   [✓] Phase 2  — Core readiness logic
   [✓] Phase 3  — Report renderer
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
   8.  Report Renderer      [✓ Phase 3]
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
  INFLUENZA_REQ:    'Influenza Req',   INFLUENZA_DUE:  'Influenza Due',   INFLUENZA_DEF:  'Influenza Deferred',
  TDAP_REQ:         'Tet/Dipth Req',   TDAP_DUE:       'Tet/Dipth Due',   TDAP_DEF:       'Tet/Dipth Deferred',
  TYPHOID_REQ:      'Typhoid Req',     TYPHOID_DUE:    'Typhoid Due',     TYPHOID_DEF:    'Typhoid Deferred',
  VARICELLA_REQ:    'Varicella Req',   VARICELLA_DUE:  'Varicella Due',   VARICELLA_DEF:  'Varicella Deferred',
  MMR_REQ:          'MMR Req',         MMR_DUE:        'MMR Due',         MMR_DEF:        'MMR Deferred',
  HEPA_REQ:         'HepA Req',        HEPA_DUE:       'HepA Due',        HEPA_DEF:       'HepA Deferred',
  HEPB_REQ:         'HepB Req',        HEPB_DUE:       'HepB Due',        HEPB_DEF:       'HepB Deferred',
  TWINRIX_REQ:      'TwinRix Req',     TWINRIX_DUE:    'TwinRix Due',     TWINRIX_DEF:    'TwinRix Deferred',
  RABIES_REQ:       'Rabies Req',      RABIES_DUE:     'Rabies Due',      RABIES_DEF:     'Rabies Deferred',
  RABIES_TITER_DUE: 'Rabies Titer Due',
  CHOLERA_REQ:      'Cholera Req',     CHOLERA_DUE:    'Cholera Due',     CHOLERA_DEF:    'Cholera Deferred',
  JEV_REQ:          'JEV Req',         JEV_DUE:        'JEV Due',         JEV_DEF:        'JEV Deferred',
  MGC_REQ:          'MGC Req',         MGC_DUE:        'MGC Due',         MGC_DEF:        'MGC Deferred',
  POLIO_REQ:        'Polio Req',       POLIO_DUE:      'Polio Due',       POLIO_DEF:      'Polio Deferred',
  YF_REQ:           'Yellow Fever Req',YF_DUE:         'Yellow Fever Due',YF_DEF:         'Yellow Fever Deferred',
  ANTHRAX_REQ:      'Anthrax Req',     ANTHRAX_DUE:    'Anthrax Due',     ANTHRAX_DEF:    'Anthrax Deferred',
  SMALLPOX_REQ:     'Smallpox Req',    SMALLPOX_DUE:   'Smallpox Due',    SMALLPOX_DEF:   'Smallpox Deferred',
  ADENOVIRUS_REQ:   'Adenovirus Req',  ADENOVIRUS_DUE: 'Adenovirus Due',  ADENOVIRUS_DEF: 'Adenovirus Deferred',
  PNEUMO_REQ:       'Pneumococcal Req',PNEUMO_DUE:     'Pneumococcal Due',PNEUMO_DEF:     'Pneumococcal Deferred',
  HPV_REQ:          'HPV Req',         HPV_DUE:        'HPV Due',         HPV_DEF:        'HPV Deferred',
};

const STATUS = {
  NA:       'NA',
  OK:       'OK',
  UPCOMING: 'UPCOMING',
  DUE_SOON: 'DUE_SOON',
  OVERDUE:  'OVERDUE',
};

const SEVERITY = { NA: 0, OK: 1, UPCOMING: 2, DUE_SOON: 3, OVERDUE: 4 };

const DEFAULT_THRESHOLDS = { yellow: 7, green: 30 };

const STORAGE_KEY     = 'hitlistHatcher_settings_v3';
const COL_WARN_YELLOW = 12;
const COL_WARN_RED    = 14;
const DEBOUNCE_MS     = 500;

const IMMUNIZATION_KEYS = [
  'influenza','tdap','typhoid','varicella','mmr',
  'hepa','hepb','twinrix','rabies','rabiesTiter',
  'cholera','jev','mgc','polio','yellowFever',
  'anthrax','smallpox','adenovirus','pneumo','hpv',
];

const IMMUNIZATION_DEFAULTS = {
  influenza:true, tdap:true, typhoid:false, varicella:false, mmr:false,
  hepa:false, hepb:false, twinrix:false, rabies:false, rabiesTiter:false,
  cholera:false, jev:false, mgc:false, polio:false, yellowFever:false,
  anthrax:false, smallpox:false, adenovirus:false, pneumo:false, hpv:false,
};

const IMMUNIZATION_LABELS = {
  influenza:'Influenza', tdap:'TDap', typhoid:'Typhoid', varicella:'Varicella',
  mmr:'MMR', hepa:'Hep A', hepb:'Hep B', twinrix:'TwinRix', rabies:'Rabies',
  rabiesTiter:'Rabies Titer', cholera:'Cholera', jev:'JEV', mgc:'MGC',
  polio:'Polio', yellowFever:'Yellow Fever', anthrax:'Anthrax',
  smallpox:'Smallpox', adenovirus:'Adenovirus', pneumo:'Pneumococcal', hpv:'HPV',
};


/* ── 2. STATE ──────────────────────────────────────────────── */

const state = {
  rawData:          null,
  personnel:        [],
  colIndex:         {},
  parseStats:       {},
  filteredResults:  [],
  currentStats:     {},
  settings:         {},
  reportGenerated:  false,
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
      const parsed  = parseMRRS(workbook);
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
  const rows  = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', raw:false });

  const headerRow = rows[MRRS_HEADER_ROW] || [];
  const colIndex  = {};
  headerRow.forEach((val, i) => {
    if (val !== null && val !== undefined && val !== '') colIndex[String(val).trim()] = i;
  });

  const missingCols = Object.values(COL).filter(c => !(c in colIndex));
  if (missingCols.length) console.warn(`parseMRRS: missing columns:`, missingCols);

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
      const t = val.trim();
      if (!t) return null;
      const iso = Date.parse(t);
      if (!isNaN(iso)) return new Date(iso);
      const mdy = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (mdy) return new Date(parseInt(mdy[3]), parseInt(mdy[1])-1, parseInt(mdy[2]));
      const dmy = t.match(/^(\d{1,2})-([A-Z]{3})-(\d{4})$/i);
      if (dmy) {
        const months={JAN:0,FEB:1,MAR:2,APR:3,MAY:4,JUN:5,JUL:6,AUG:7,SEP:8,OCT:9,NOV:10,DEC:11};
        const m = months[dmy[2].toUpperCase()];
        if (m !== undefined) return new Date(parseInt(dmy[3]), m, parseInt(dmy[1]));
      }
    }
    return null;
  }

  function dentalDueDate(row, useMrrsDate=false) {
    if (useMrrsDate) return parseDate(getCell(row,'DENTAL_DUE'));
    const examDt = parseDate(getCell(row,'DENTAL_DT'));
    if (!examDt) return null;
    const due = new Date(examDt);
    due.setDate(due.getDate() + 365);
    return due;
  }

  function parseDentalClass(val) {
    if (!val || String(val).trim()==='') return 4;
    const n = parseInt(String(val).trim(), 10);
    return (n>=1 && n<=4) ? n : 4;
  }

  function isYes(val) {
    if (!val && val!==0) return false;
    const s = String(val).trim().toLowerCase();
    return s==='yes'||s==='y'||s==='1'||s==='true';
  }

  function isDeferred(val) {
    if (!val && val!==0) return false;
    const s = String(val).trim().toLowerCase();
    return s==='yes'||s==='y'||s==='1'||s==='true'||s==='deferred';
  }

  function sectionLabel(row) {
    const dept    = String(getCell(row,'COMP_DEPT')).trim();
    const platoon = String(getCell(row,'PLATOON')).trim();
    if (dept && platoon) return `${dept}-${platoon}`;
    return dept || platoon || '';
  }

  const personnel=[], imm=(k,rq,du,df,lb)=>({
    req:isYes(getCell(row,rq)), due:parseDate(getCell(row,du)),
    deferred:isDeferred(getCell(row,df)), label:lb
  });
  let officers=0, enlisted=0, skipped=0;

  for (let i=MRRS_DATA_START; i<rows.length; i++) {
    const row  = rows[i];
    const name = String(getCell(row,'NAME')).trim();
    if (!name) { skipped++; continue; }

    const offEnl    = String(getCell(row,'OFF_ENL')).trim();
    const rank      = String(getCell(row,'RANK')).trim();
    const isOfficer = offEnl.toLowerCase().includes('officer') ||
                      (offEnl==='' && rank.toUpperCase().includes('CWO'));

    if (isOfficer) officers++; else enlisted++;

    personnel.push({
      name, rank,
      section:      sectionLabel(row),
      offEnl:       isOfficer ? 'Officer' : 'Enlisted',
      sex:          String(getCell(row,'SEX')).trim(),
      imrStatus:    String(getCell(row,'IMR_STATUS')).trim(),
      deploying:    isYes(getCell(row,'DEPLOYING')),
      phaDue:       parseDate(getCell(row,'PHA_DUE')),
      dentalDue:    dentalDueDate(row, false),
      dentalClass:  parseDentalClass(getCell(row,'DENTAL_CLASS')),
      hivDue:       parseDate(getCell(row,'HIV_DUE')),
      audioDue:     parseDate(getCell(row,'AUDIO_DUE')),
      pdhaDue:      parseDate(getCell(row,'PDHA_DUE')),
      pdhraDue:     parseDate(getCell(row,'PDHRA_DUE')),
      tstDue:       parseDate(getCell(row,'TST_DUE')),
      tstQuestDue:  parseDate(getCell(row,'TST_QUEST_DUE')),
      mammogramDue: parseDate(getCell(row,'MAMMOGRAM_DUE')),
      papDue:       parseDate(getCell(row,'PAP_DUE')),
      immunizations: {
        influenza:   {req:isYes(getCell(row,'INFLUENZA_REQ')),  due:parseDate(getCell(row,'INFLUENZA_DUE')),  deferred:isDeferred(getCell(row,'INFLUENZA_DEF')),  label:'Influenza'},
        tdap:        {req:isYes(getCell(row,'TDAP_REQ')),       due:parseDate(getCell(row,'TDAP_DUE')),       deferred:isDeferred(getCell(row,'TDAP_DEF')),       label:'TDap'},
        typhoid:     {req:isYes(getCell(row,'TYPHOID_REQ')),    due:parseDate(getCell(row,'TYPHOID_DUE')),    deferred:isDeferred(getCell(row,'TYPHOID_DEF')),    label:'Typhoid'},
        varicella:   {req:isYes(getCell(row,'VARICELLA_REQ')), due:parseDate(getCell(row,'VARICELLA_DUE')), deferred:isDeferred(getCell(row,'VARICELLA_DEF')), label:'Varicella'},
        mmr:         {req:isYes(getCell(row,'MMR_REQ')),        due:parseDate(getCell(row,'MMR_DUE')),        deferred:isDeferred(getCell(row,'MMR_DEF')),        label:'MMR'},
        hepa:        {req:isYes(getCell(row,'HEPA_REQ')),       due:parseDate(getCell(row,'HEPA_DUE')),       deferred:isDeferred(getCell(row,'HEPA_DEF')),       label:'Hep A'},
        hepb:        {req:isYes(getCell(row,'HEPB_REQ')),       due:parseDate(getCell(row,'HEPB_DUE')),       deferred:isDeferred(getCell(row,'HEPB_DEF')),       label:'Hep B'},
        twinrix:     {req:isYes(getCell(row,'TWINRIX_REQ')),    due:parseDate(getCell(row,'TWINRIX_DUE')),    deferred:isDeferred(getCell(row,'TWINRIX_DEF')),    label:'TwinRix'},
        rabies:      {req:isYes(getCell(row,'RABIES_REQ')),     due:parseDate(getCell(row,'RABIES_DUE')),     deferred:isDeferred(getCell(row,'RABIES_DEF')),     label:'Rabies'},
        rabiesTiter: {req:isYes(getCell(row,'RABIES_REQ')),     due:parseDate(getCell(row,'RABIES_TITER_DUE')),deferred:false,                                    label:'Rabies Titer'},
        cholera:     {req:isYes(getCell(row,'CHOLERA_REQ')),    due:parseDate(getCell(row,'CHOLERA_DUE')),    deferred:isDeferred(getCell(row,'CHOLERA_DEF')),    label:'Cholera'},
        jev:         {req:isYes(getCell(row,'JEV_REQ')),        due:parseDate(getCell(row,'JEV_DUE')),        deferred:isDeferred(getCell(row,'JEV_DEF')),        label:'JEV'},
        mgc:         {req:isYes(getCell(row,'MGC_REQ')),        due:parseDate(getCell(row,'MGC_DUE')),        deferred:isDeferred(getCell(row,'MGC_DEF')),        label:'MGC'},
        polio:       {req:isYes(getCell(row,'POLIO_REQ')),      due:parseDate(getCell(row,'POLIO_DUE')),      deferred:isDeferred(getCell(row,'POLIO_DEF')),      label:'Polio'},
        yellowFever: {req:isYes(getCell(row,'YF_REQ')),         due:parseDate(getCell(row,'YF_DUE')),         deferred:isDeferred(getCell(row,'YF_DEF')),         label:'Yellow Fever'},
        anthrax:     {req:isYes(getCell(row,'ANTHRAX_REQ')),    due:parseDate(getCell(row,'ANTHRAX_DUE')),    deferred:isDeferred(getCell(row,'ANTHRAX_DEF')),    label:'Anthrax'},
        smallpox:    {req:isYes(getCell(row,'SMALLPOX_REQ')),   due:parseDate(getCell(row,'SMALLPOX_DUE')),   deferred:isDeferred(getCell(row,'SMALLPOX_DEF')),   label:'Smallpox'},
        adenovirus:  {req:isYes(getCell(row,'ADENOVIRUS_REQ')), due:parseDate(getCell(row,'ADENOVIRUS_DUE')), deferred:isDeferred(getCell(row,'ADENOVIRUS_DEF')), label:'Adenovirus'},
        pneumo:      {req:isYes(getCell(row,'PNEUMO_REQ')),     due:parseDate(getCell(row,'PNEUMO_DUE')),     deferred:isDeferred(getCell(row,'PNEUMO_DEF')),     label:'Pneumococcal'},
        hpv:         {req:isYes(getCell(row,'HPV_REQ')),        due:parseDate(getCell(row,'HPV_DUE')),        deferred:isDeferred(getCell(row,'HPV_DEF')),        label:'HPV'},
      },
    });
  }

  return {
    personnel, colIndex,
    stats:{ total:personnel.length, officers, enlisted, skipped, columns:Object.keys(colIndex).length }
  };
}

function logParseResults(parsed) {
  const { personnel, stats } = parsed;
  console.group('%cHitlist Hatcher — Parse Results','color:#003087;font-weight:bold;font-size:14px;');
  console.table({'Total':stats.total,'Officers':stats.officers,'Enlisted':stats.enlisted,'Skipped':stats.skipped,'Columns':stats.columns});
  console.table(personnel.slice(0,5).map(p=>({name:p.name,rank:p.rank,section:p.section,category:p.offEnl,dentalClass:p.dentalClass})));
  console.table(personnel.slice(0,5).map(p=>({name:p.name,phaDue:p.phaDue?formatDateFull(p.phaDue):'null',dentalDue:p.dentalDue?formatDateFull(p.dentalDue):'null',hivDue:p.hivDue?formatDateFull(p.hivDue):'null'})));
  window._hhPersonnel = personnel;
  console.log('%cFull array: window._hhPersonnel','color:#1e8449;font-style:italic;');
  console.groupEnd();
}


/* ── 7. READINESS LOGIC ────────────────────────────────────── */

function evaluateItem(dueDate, projectionDate, thresholds) {
  if (!dueDate || !(dueDate instanceof Date) || isNaN(dueDate)) return STATUS.NA;
  const days = daysDiff(dueDate, projectionDate);
  if (days <= 0)                 return STATUS.OVERDUE;
  if (days <= thresholds.yellow) return STATUS.DUE_SOON;
  if (days <= thresholds.green)  return STATUS.UPCOMING;
  return STATUS.OK;
}

function evaluateDental(person, projectionDate, thresholds) {
  if (person.dentalClass === 3) return { status: STATUS.OVERDUE,  displayText: 'Class 3' };
  if (person.dentalClass === 4) return { status: STATUS.OVERDUE,  displayText: 'Class 4' };
  return { status: evaluateItem(person.dentalDue, projectionDate, thresholds), displayText: null };
}

function evaluateImmunization(immObj, projectionDate, thresholds) {
  if (!immObj.req)     return STATUS.NA;
  if (immObj.deferred) return STATUS.NA;
  if (!immObj.due)     return STATUS.OVERDUE;
  return evaluateItem(immObj.due, projectionDate, thresholds);
}

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

function evaluatePerson(person, settings, projectionDate) {
  const t = settings.thresholds || DEFAULT_THRESHOLDS;

  const phaResult   = settings.items.pha
    ? { status: evaluateItem(person.phaDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  const dentalResult = settings.items.dental
    ? evaluateDental(person, projectionDate, t)
    : { status: STATUS.NA, displayText: null };

  const hivResult   = settings.items.hiv
    ? { status: evaluateItem(person.hivDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  const audioResult = settings.items.audio
    ? { status: evaluateItem(person.audioDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  const pdhaResult  = settings.items.pdha
    ? { status: evaluateItem(person.pdhaDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  const pdhraResult = settings.items.pdhra
    ? { status: evaluateItem(person.pdhraDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  const tstResult   = settings.items.tst
    ? { status: evaluateItem(person.tstDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  const tstQuestResult = settings.items.tstQuest
    ? { status: evaluateItem(person.tstQuestDue, projectionDate, t), displayText: null }
    : { status: STATUS.NA, displayText: null };

  let wellWomanResult = { status: STATUS.NA, displayText: null };
  if (settings.items.wellWoman && person.sex.toLowerCase() === 'female') {
    const mamStatus = evaluateItem(person.mammogramDue, projectionDate, t);
    const papStatus = evaluateItem(person.papDue, projectionDate, t);
    wellWomanResult = {
      status: SEVERITY[mamStatus] >= SEVERITY[papStatus] ? mamStatus : papStatus,
      displayText: null,
    };
  }

  const immKeys    = settings.immunizationKeys || [];
  const immResult  = evaluateImmunizations(person, immKeys, projectionDate, t);
  const immunizationsResult = {
    grouped:    { status: immResult.worstStatus, count: immResult.count, dueNames: immResult.dueNames },
    perVaccine: immResult.perVaccine,
  };

  const allStatuses = [
    phaResult.status, dentalResult.status, hivResult.status, audioResult.status,
    pdhaResult.status, pdhraResult.status, tstResult.status, tstQuestResult.status,
    wellWomanResult.status, immResult.worstStatus,
  ];
  const overallStatus = allStatuses.reduce(
    (worst,s) => SEVERITY[s] > SEVERITY[worst] ? s : worst, STATUS.NA
  );
  const isDue = SEVERITY[overallStatus] >= SEVERITY[STATUS.UPCOMING];

  return {
    person, overallStatus, isDue,
    items: { pha:phaResult, dental:dentalResult, hiv:hivResult, audio:audioResult,
             pdha:pdhaResult, pdhra:pdhraResult, tst:tstResult, tstQuest:tstQuestResult,
             wellWoman:wellWomanResult, immunizations:immunizationsResult },
  };
}

function applyFilters(personnel, settings, projectionDate) {
  const allEvaluated = personnel.map(p => evaluatePerson(p, settings, projectionDate));
  const stats        = computeReadinessStats(personnel);
  const offEnlFilter = settings.offEnlFilter || 'combined';

  let filtered = allEvaluated.filter(r => {
    if (!r.isDue) return false;
    if (offEnlFilter === 'officer')  return r.person.offEnl === 'Officer';
    if (offEnlFilter === 'enlisted') return r.person.offEnl === 'Enlisted';
    return true;
  });

  const sortBy = settings.sortBy || 'name';
  filtered.sort((a, b) => {
    if (sortBy === 'section') {
      const cmp = a.person.section.localeCompare(b.person.section);
      if (cmp !== 0) return cmp;
    }
    return a.person.name.localeCompare(b.person.name);
  });

  return { results: filtered, stats };
}

function computeReadinessStats(personnel) {
  let fullyReady=0, notReady=0, partial=0, indeterminate=0;
  personnel.forEach(p => {
    const s = (p.imrStatus||'').toLowerCase();
    if      (s.includes('fully'))   fullyReady++;
    else if (s.includes('not'))     notReady++;
    else if (s.includes('partial')) partial++;
    else                             indeterminate++;
  });
  const total = personnel.length;
  const pct   = n => total>0 ? ((n/total)*100).toFixed(1) : '0.0';
  return {
    total, fullyReady, notReady, partial, indeterminate,
    fullyReadyPct:pct(fullyReady), notReadyPct:pct(notReady),
    partialPct:pct(partial), indeterminatePct:pct(indeterminate),
  };
}


/* ── 8. REPORT RENDERER ────────────────────────────────────── */

/*
  8a. renderReport(results, stats, settings, projectionDate)
  ──────────────────────────────────────────────────────────
  Main entry point. Builds the complete hit list HTML and injects
  it into dom.reportOutput. Also shows the export bar.
*/
function renderReport(results, stats, settings, projectionDate) {
  if (results.length === 0) {
    dom.previewPlaceholder.hidden = true;
    dom.reportOutput.innerHTML = `
      <div style="background:#fff;border-radius:6px;padding:48px;text-align:center;box-shadow:0 2px 10px rgba(0,0,0,.1);">
        <div style="font-size:48px;margin-bottom:16px;">✓</div>
        <p style="font-size:18px;font-weight:700;color:#003087;margin-bottom:8px;">No personnel due</p>
        <p style="font-size:14px;color:#444;">
          No one is due or upcoming for the selected items within the current warning windows.
        </p>
      </div>`;
    dom.exportBar.hidden = true;
    return;
  }

  const colDefs = getColumnDefs(settings);
  updateColumnCounter(colDefs.length);

  const html = `
    <div class="hitlist-wrapper">
      ${buildHeader(stats, settings, projectionDate)}
      ${buildLegend(settings)}
      <div class="hitlist-table-wrapper">
        ${buildTable(results, colDefs, settings, projectionDate)}
      </div>
    </div>`;

  dom.previewPlaceholder.hidden = true;
  dom.reportOutput.innerHTML    = html;
  dom.exportBar.hidden          = false;
  state.reportGenerated         = true;
}


/*
  8b. getColumnDefs(settings)
  ───────────────────────────
  Returns an ordered array of column definition objects.
  Each object drives both the <th> header and the <td> data cell.

  { key, label, type }
    key   — matches result.items property name, or special values
             'name' | 'rank' | 'section' | 'imm-grouped' | 'imm-{vaccineKey}'
    label — displayed in the column header
    type  — 'identity' | 'item' | 'imm-grouped' | 'imm-individual'
*/
function getColumnDefs(settings) {
  const defs = [];

  // Identity columns — always first
  defs.push({ key:'name',    label:'Name',    type:'identity' });
  if (settings.showRank)    defs.push({ key:'rank',    label:'Rank',    type:'identity' });
  if (settings.showSection) defs.push({ key:'section', label:'Section', type:'identity' });

  // Core readiness items
  if (settings.items.pha)       defs.push({ key:'pha',       label:'PHA',        type:'item' });
  if (settings.items.dental)    defs.push({ key:'dental',    label:'Dental',     type:'item' });
  if (settings.items.hiv)       defs.push({ key:'hiv',       label:'HIV Lab',    type:'item' });
  if (settings.items.audio)     defs.push({ key:'audio',     label:'Audiogram',  type:'item' });

  // Deployment (default off)
  if (settings.items.pdha)      defs.push({ key:'pdha',      label:'PDHA',       type:'item' });
  if (settings.items.pdhra)     defs.push({ key:'pdhra',     label:'PDHRA',      type:'item' });

  // Other
  if (settings.items.tst)       defs.push({ key:'tst',       label:'TST',        type:'item' });
  if (settings.items.tstQuest)  defs.push({ key:'tstQuest',  label:'TST Quest',  type:'item' });
  if (settings.items.wellWoman) defs.push({ key:'wellWoman', label:'Well-Woman', type:'item' });

  // Immunizations
  if (settings.items.immunizations) {
    const mode = settings.immunDisplayMode || 'grouped';
    if (mode === 'grouped') {
      defs.push({ key:'imm-grouped', label:'Immunizations', type:'imm-grouped' });
    } else {
      // Individual column per selected vaccine
      (settings.immunizationKeys || []).forEach(vaccKey => {
        defs.push({
          key:   `imm-${vaccKey}`,
          label: IMMUNIZATION_LABELS[vaccKey] || vaccKey,
          type:  'imm-individual',
          vaccKey,
        });
      });
    }
  }

  return defs;
}


/*
  8c. buildHeader(stats, settings, projectionDate)
  ─────────────────────────────────────────────────
  Returns HTML string for the full report header block:
  emblems, unit name, title, projection date, readiness stats,
  and custom clinic info text.
*/
function buildHeader(stats, settings, projectionDate) {
  const unitName    = escHtml(settings.unitName || 'Unit Name');
  const customText  = escHtml(settings.customText || '');
  const projStr     = formatDateFull(projectionDate);
  const reportDate  = formatDateFull(new Date());

  const emblemLeft  = settings.emblemBase64
    ? `<img class="hitlist-emblem" src="${settings.emblemBase64}" alt="Unit Emblem" />`
    : `<div class="hitlist-emblem-placeholder">Unit<br/>Emblem</div>`;

  const emblemRight = settings.emblemBase64
    ? `<img class="hitlist-emblem" src="${settings.emblemBase64}" alt="Unit Emblem" />`
    : `<div class="hitlist-emblem-placeholder">Unit<br/>Emblem</div>`;

  const infoBar = customText
    ? `<div class="hitlist-info-bar">${customText}</div>`
    : '';

  return `
    <div class="hitlist-header">
      <div class="hitlist-header-top">
        ${emblemLeft}
        <div class="hitlist-title-block">
          <div class="hitlist-unit-name">${unitName} Medical Hit List</div>
          <div class="hitlist-title-line">Report Date: ${reportDate}</div>
          <div class="hitlist-projection-line">Projected to: ${projStr}</div>
        </div>
        ${emblemRight}
      </div>
      <div class="hitlist-stats-bar">
        <div class="hitlist-stat">
          <span class="hitlist-stat-label">Fully Ready</span>
          <span class="hitlist-stat-value stat-good">${stats.fullyReadyPct}%</span>
        </div>
        <div class="hitlist-stat">
          <span class="hitlist-stat-label">Not Ready</span>
          <span class="hitlist-stat-value stat-bad">${stats.notReadyPct}%</span>
        </div>
        <div class="hitlist-stat">
          <span class="hitlist-stat-label">Partial</span>
          <span class="hitlist-stat-value stat-warn">${stats.partialPct}%</span>
        </div>
        <div class="hitlist-stat">
          <span class="hitlist-stat-label">Indeterminate</span>
          <span class="hitlist-stat-value stat-neutral">${stats.indeterminatePct}%</span>
        </div>
      </div>
      ${infoBar}
    </div>`;
}


/*
  8d. buildLegend(settings)
  ─────────────────────────
  Returns the color key legend HTML.
*/
function buildLegend(settings) {
  const t = settings.thresholds || DEFAULT_THRESHOLDS;
  return `
    <div class="hitlist-legend">
      <span class="hitlist-legend-label">Key:</span>
      <span class="legend-item"><span class="legend-swatch red"></span>Overdue / Class 3-4</span>
      <span class="legend-item"><span class="legend-swatch yellow"></span>Due within ${t.yellow} days</span>
      <span class="legend-item"><span class="legend-swatch green"></span>Due within ${t.green} days</span>
      <span class="legend-item"><span class="legend-swatch none" style="border:1px solid #ccc;"></span>Not due</span>
    </div>`;
}


/*
  8e. buildTable(results, colDefs, settings, projectionDate)
  ───────────────────────────────────────────────────────────
  Returns the full <table> HTML with header row and all data rows.
*/
function buildTable(results, colDefs, settings, projectionDate) {
  const headerCells = colDefs.map(col => {
    const cls = col.type === 'identity' && col.key === 'name' ? ' col-name-hdr' : '';
    return `<th class="${cls}">${escHtml(col.label)}</th>`;
  }).join('');

  const dataRows = results.map(result => buildRow(result, colDefs, settings, projectionDate)).join('');

  return `
    <table class="hitlist-table">
      <thead><tr>${headerCells}</tr></thead>
      <tbody>${dataRows}</tbody>
    </table>`;
}


/*
  8f. buildRow(result, colDefs, settings, projectionDate)
  ────────────────────────────────────────────────────────
  Returns HTML for one <tr> covering a single personnel record.
*/
function buildRow(result, colDefs, settings, projectionDate) {
  const cells = colDefs.map(col => getCellHTML(result, col, settings, projectionDate)).join('');
  return `<tr>${cells}</tr>`;
}


/*
  8g. getCellHTML(result, col, settings, projectionDate)
  ───────────────────────────────────────────────────────
  Returns HTML for one <td> cell.

  Handles:
  - Identity fields (name, rank, section) — plain text
  - Standard item fields (pha, hiv, etc.) — colored by status, shows date
  - Dental — colored by status, shows displayText (Class 3/4) or date
  - Immunizations grouped — count display with tooltip
  - Immunizations individual — per-vaccine colored cell
*/
function getCellHTML(result, col, settings, projectionDate) {
  const { key, type, vaccKey } = col;

  // ── Identity fields ─────────────────────────────────────────
  if (type === 'identity') {
    if (key === 'name')    return `<td class="col-name">${escHtml(result.person.name)}</td>`;
    if (key === 'rank')    return `<td class="col-rank">${escHtml(result.person.rank)}</td>`;
    if (key === 'section') return `<td class="col-section">${escHtml(result.person.section)}</td>`;
  }

  // ── Grouped immunizations ────────────────────────────────────
  if (type === 'imm-grouped') {
    const grp = result.items.immunizations.grouped;
    if (grp.status === STATUS.NA || grp.count === 0) {
      return `<td class="col-item cell-na">—</td>`;
    }
    const statusClass = statusToCssClass(grp.status);
    const countLabel  = `${grp.count} due`;
    const tooltipLines = grp.dueNames.map(n => escHtml(n)).join('<br/>');
    return `
      <td class="col-item ${statusClass}">
        <div class="imm-grouped">
          ${escHtml(countLabel)}
          <div class="imm-tooltip">${tooltipLines}</div>
        </div>
      </td>`;
  }

  // ── Individual immunization column ───────────────────────────
  if (type === 'imm-individual') {
    const pv = result.items.immunizations.perVaccine[vaccKey];
    if (!pv || pv.status === STATUS.NA || pv.status === STATUS.OK) {
      return `<td class="col-item cell-na">—</td>`;
    }
    const statusClass = statusToCssClass(pv.status);
    const dateStr     = pv.due ? formatDate(pv.due) : '?';
    return `<td class="col-item ${statusClass}">${dateStr}</td>`;
  }

  // ── Standard readiness item ──────────────────────────────────
  if (type === 'item') {
    const itemResult = result.items[key];
    if (!itemResult || itemResult.status === STATUS.NA || itemResult.status === STATUS.OK) {
      return `<td class="col-item cell-na">—</td>`;
    }

    const statusClass = statusToCssClass(itemResult.status);

    // Dental: may have a displayText override (Class 3 / Class 4)
    if (key === 'dental' && itemResult.displayText) {
      return `<td class="col-item ${statusClass}">${escHtml(itemResult.displayText)}</td>`;
    }

    // All other items: show formatted due date
    const dueDate = getDueDateForItem(result.person, key);
    const dateStr = dueDate ? formatDate(dueDate) : '?';
    return `<td class="col-item ${statusClass}">${dateStr}</td>`;
  }

  // Fallback — should not reach here
  return `<td class="col-item cell-na">—</td>`;
}


/*
  8h. getDueDateForItem(person, itemKey)
  ───────────────────────────────────────
  Maps a result item key back to the corresponding due date on
  the person object so the renderer can format it for display.
*/
function getDueDateForItem(person, itemKey) {
  const map = {
    pha:       person.phaDue,
    dental:    person.dentalDue,
    hiv:       person.hivDue,
    audio:     person.audioDue,
    pdha:      person.pdhaDue,
    pdhra:     person.pdhraDue,
    tst:       person.tstDue,
    tstQuest:  person.tstQuestDue,
    wellWoman: person.mammogramDue || person.papDue,
  };
  return map[itemKey] || null;
}


/*
  8i. statusToCssClass(status)
  ─────────────────────────────
  Maps a STATUS constant to the corresponding CSS class.
*/
function statusToCssClass(status) {
  if (status === STATUS.OVERDUE)  return 'status-red';
  if (status === STATUS.DUE_SOON) return 'status-yellow';
  if (status === STATUS.UPCOMING) return 'status-green';
  return '';
}


/*
  8j. escHtml(str)
  ─────────────────
  Escapes a string for safe insertion into HTML.
  Prevents XSS from any unexpected MRRS data values.
*/
function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


/* ── 9. SETTINGS PANEL ─────────────────────────────────────── */
/*  Built in Phase 4  */

function initSettingsPanel() {
  console.log('initSettingsPanel() — Phase 4 not yet built.');
}


/*
  Generate button — now chains Phases 1, 2, and 3 together.
  Uses temporary hardcoded settings until Phase 4 builds the UI.
*/
dom.generateBtn.addEventListener('click', () => {
  if (!state.rawData || state.personnel.length === 0) {
    alert('Please upload an MRRS file first.');
    return;
  }

  // ── Temporary settings — replaced entirely by Phase 4 UI ────
  const tempSettings = {
    unitName:          'VMFA-314',          // unit name in header
    customText:        'PHA & Shots (FLAS): Mon–Fri 0700–1100\nHIV Lab: Mon–Fri 0730–1500',
    emblemBase64:      null,                // no emblem until Phase 4 upload
    items: {
      pha:            true,
      dental:         true,
      hiv:            true,
      audio:          true,
      pdha:           false,
      pdhra:          false,
      tst:            false,
      tstQuest:       false,
      wellWoman:      false,
      immunizations:  true,
    },
    immunizationKeys:  ['influenza', 'tdap'],
    immunDisplayMode:  'grouped',           // 'grouped' | 'individual'
    thresholds:        { yellow: 7, green: 30 },
    offEnlFilter:      'combined',          // 'combined' | 'officer' | 'enlisted'
    sortBy:            'name',              // 'name' | 'section'
    showRank:          true,
    showSection:       true,
    dentalUseMrrsDate: false,
  };

  const projectionDate = new Date();

  const { results, stats } = applyFilters(state.personnel, tempSettings, projectionDate);
  state.filteredResults = results;
  state.currentStats    = stats;

  console.log(`Generate: ${results.length} personnel flagged. Rendering report...`);
  renderReport(results, stats, tempSettings, projectionDate);
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
  if (bytes < 1048576)     return `${(bytes/1024).toFixed(1)} KB`;
  return `${(bytes/1048576).toFixed(1)} MB`;
}

function formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  const m = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${date.getDate()} ${m[date.getMonth()]}`;
}

function formatDateFull(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return '';
  const m = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return `${date.getDate()} ${m[date.getMonth()]} ${date.getFullYear()}`;
}

function daysDiff(dateA, dateB) {
  return Math.round((dateA - dateB) / 86400000);
}

function debounce(fn, delay) {
  let timer;
  return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), delay); };
}

function updateColumnCounter(count) {
  dom.columnCount.textContent = count;
  dom.columnCounter.classList.remove('warn-yellow','warn-red');
  if      (count >= COL_WARN_RED)    dom.columnCounter.classList.add('warn-red');
  else if (count >= COL_WARN_YELLOW) dom.columnCounter.classList.add('warn-yellow');
}
