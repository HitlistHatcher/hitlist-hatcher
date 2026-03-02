/* ============================================================
   HITLIST HATCHER 3.0 — APPLICATION LOGIC  (Phase 4b)
============================================================ */

/* ── 1. CONSTANTS ──────────────────────────────────────────── */

const APP_VERSION = '3.0.0';
const MRRS_HEADER_ROW = 2;
const MRRS_DATA_START = 3;
const MRRS_SHEET_NAME = 'IMR Detail';

const COL = {
  NAME:'Name', RANK:'Rank/Rate', COMP_DEPT:'Comp/Dept', PLATOON:'Platoon',
  OFF_ENL:'Off Enl Indicator', SEX:'Sex', IMR_STATUS:'IMR Status', DEPLOYING:'Deploying',
  PHA_DT:'PHA Dt', PHA_DUE:'PHA Due',
  DENTAL_DT:'Dental Exam Dt', DENTAL_DUE:'Dental Exam Due', DENTAL_CLASS:'Dental Cond Code',
  HIV_DT:'HIV Test Dt', HIV_DUE:'HIV Test Due',
  AUDIO_DT:'Audio 2216 Dt', AUDIO_DUE:'Audio 2216 Due',
  PDHA_DUE:'PDHA Due', PDHRA_DUE:'PDHRA Due',
  MHA2_STATUS:'MHA2 Status', MHA3_STATUS:'MHA3 Status', MHA4_STATUS:'MHA4 Status',
  WARNING_TAG_DUE:'Warning Tag Due Dt',
  VERIFY_GLASSES_DUE:'Two Pair Verified Due',
  VERIFY_INSERTS_DUE:'Inserts Due',
  BLOOD_TYPE:'Blood Type',
  REF_AUDIO_DUE:'DD2215 Due',
  DNA_DUE:'DNA Due Dt',
  G6PD_DUE:'G6PD Due Dt',
  SICKLE_DUE:'Sickle Cell Due Dt',
  TST_DUE:'TST Due Dt', TST_QUEST_DUE:'TST Quest Due',
  MAMMOGRAM_DUE:'Mammogram Due', PAP_DUE:'Pap Smear Due',
  INFLUENZA_REQ:'Influenza Req',   INFLUENZA_DUE:'Influenza Due',   INFLUENZA_DEF:'Influenza Deferred',
  TDAP_REQ:'Tet/Dipth Req',        TDAP_DUE:'Tet/Dipth Due',        TDAP_DEF:'Tet/Dipth Deferred',
  TYPHOID_REQ:'Typhoid Req',       TYPHOID_DUE:'Typhoid Due',       TYPHOID_DEF:'Typhoid Deferred',
  VARICELLA_REQ:'Varicella Req',   VARICELLA_DUE:'Varicella Due',   VARICELLA_DEF:'Varicella Deferred',
  MMR_REQ:'MMR Req',               MMR_DUE:'MMR Due',               MMR_DEF:'MMR Deferred',
  HEPA_REQ:'HepA Req',             HEPA_DUE:'HepA Due',             HEPA_DEF:'HepA Deferred',
  HEPB_REQ:'HepB Req',             HEPB_DUE:'HepB Due',             HEPB_DEF:'HepB Deferred',
  TWINRIX_REQ:'TwinRix Req',       TWINRIX_DUE:'TwinRix Due',       TWINRIX_DEF:'TwinRix Deferred',
  RABIES_REQ:'Rabies Req',         RABIES_DUE:'Rabies Due',         RABIES_DEF:'Rabies Deferred',
  RABIES_TITER_DUE:'Rabies Titer Due',
  CHOLERA_REQ:'Cholera Req',       CHOLERA_DUE:'Cholera Due',       CHOLERA_DEF:'Cholera Deferred',
  JEV_REQ:'JEV Req',               JEV_DUE:'JEV Due',               JEV_DEF:'JEV Deferred',
  MGC_REQ:'MGC Req',               MGC_DUE:'MGC Due',               MGC_DEF:'MGC Deferred',
  POLIO_REQ:'Polio Req',           POLIO_DUE:'Polio Due',           POLIO_DEF:'Polio Deferred',
  YF_REQ:'Yellow Fever Req',       YF_DUE:'Yellow Fever Due',       YF_DEF:'Yellow Fever Deferred',
  ANTHRAX_REQ:'Anthrax Req',       ANTHRAX_DUE:'Anthrax Due',       ANTHRAX_DEF:'Anthrax Deferred',
  SMALLPOX_REQ:'Smallpox Req',     SMALLPOX_DUE:'Smallpox Due',     SMALLPOX_DEF:'Smallpox Deferred',
  ADENOVIRUS_REQ:'Adenovirus Req', ADENOVIRUS_DUE:'Adenovirus Due', ADENOVIRUS_DEF:'Adenovirus Deferred',
  PNEUMO_REQ:'Pneumococcal Req',   PNEUMO_DUE:'Pneumococcal Due',   PNEUMO_DEF:'Pneumococcal Deferred',
  HPV_REQ:'HPV Req',               HPV_DUE:'HPV Due',               HPV_DEF:'HPV Deferred',
};

const STATUS   = { NA:'NA', OK:'OK', UPCOMING:'UPCOMING', DUE_SOON:'DUE_SOON', OVERDUE:'OVERDUE' };
const SEVERITY = { NA:0, OK:1, UPCOMING:2, DUE_SOON:3, OVERDUE:4 };
const DEFAULT_THRESHOLDS = { yellow:7, green:30 };

const STORAGE_KEY      = 'hitlistHatcher_settings_v3';
const DISCLAIMER_KEY   = 'hitlistHatcher_disclaimerSeen';
const COL_WARN_YELLOW  = 12;
const COL_WARN_RED     = 14;
const DEBOUNCE_MS      = 400;

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
  rawData:null, personnel:[], colIndex:{}, parseStats:{},
  filteredResults:[], currentStats:{}, settings:{},
  reportGenerated:false, emblemBase64:null,
  immDisplayMode:'grouped', _lastDentalFlag:false,
};


/* ── 3. DOM REFERENCES ─────────────────────────────────────── */

const dom = {
  disclaimerModal:     document.getElementById('disclaimerModal'),
  disclaimerAccept:    document.getElementById('disclaimerAccept'),
  disclaimerRemember:  document.getElementById('disclaimerRemember'),
  uploadArea:          document.getElementById('uploadArea'),
  fileInput:           document.getElementById('fileInput'),
  uploadStatus:        document.getElementById('uploadStatus'),
  itemPha:             document.getElementById('item-pha'),
  itemDental:          document.getElementById('item-dental'),
  dentalSub:           document.getElementById('dentalSub'),
  itemHiv:             document.getElementById('item-hiv'),
  itemAudio:           document.getElementById('item-audio'),
  itemPdha:            document.getElementById('item-pdha'),
  itemPdhra:           document.getElementById('item-pdhra'),
  itemMha2:            document.getElementById('item-mha2'),
  itemMha3:            document.getElementById('item-mha3'),
  itemMha4:            document.getElementById('item-mha4'),
  itemVerifyGlasses:   document.getElementById('item-verifyGlasses'),
  itemVerifyInserts:   document.getElementById('item-verifyInserts'),
  itemWarningTag:      document.getElementById('item-warningTag'),
  itemBloodType:       document.getElementById('item-bloodType'),
  itemRefAudio:        document.getElementById('item-refAudio'),
  itemDna:             document.getElementById('item-dna'),
  itemG6pd:            document.getElementById('item-g6pd'),
  itemSickle:          document.getElementById('item-sickle'),
  itemTst:             document.getElementById('item-tst'),
  itemTstQuest:        document.getElementById('item-tstQuest'),
  itemWellWoman:       document.getElementById('item-wellWoman'),
  immModeGroupedBtn:   document.getElementById('immModeGroupedBtn'),
  immModeIndividualBtn:document.getElementById('immModeIndividualBtn'),
  immCheckboxes:       document.getElementById('immCheckboxes'),
  immSelectAll:        document.getElementById('immSelectAll'),
  immClearAll:         document.getElementById('immClearAll'),
  offEnlFilter:        document.getElementById('offEnlFilter'),
  sortBy:              document.getElementById('sortBy'),
  showRank:            document.getElementById('showRank'),
  showSection:         document.getElementById('showSection'),
  threshYellow:        document.getElementById('threshYellow'),
  threshGreen:         document.getElementById('threshGreen'),
  threshError:         document.getElementById('threshError'),
  unitName:            document.getElementById('unitName'),
  emblemInput:         document.getElementById('emblemInput'),
  clearEmblemBtn:      document.getElementById('clearEmblemBtn'),
  emblemPreview:       document.getElementById('emblemPreview'),
  infoColCount:        document.getElementById('infoColCount'),
  infoColContainer:    document.getElementById('infoColContainer'),
  projectionDate:      document.getElementById('projectionDate'),
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
  floatingTooltip:     document.getElementById('floatingTooltip'),
};


/* ── 4. INITIALIZATION ─────────────────────────────────────── */

document.addEventListener('DOMContentLoaded', () => {
  console.log(`Hitlist Hatcher ${APP_VERSION} — initializing`);
  initDisclaimer();
  initSettingsPanel();
  wireUploadHandlers();
  wireExportPdfHandler();
  wireSettingsHandlers();
  wireTooltips();
  setDefaultProjectionDate();
  refreshColumnCounter();
  console.log('Ready.');
});


/* ── 4a. DISCLAIMER ────────────────────────────────────────── */

function initDisclaimer() {
  if (localStorage.getItem(DISCLAIMER_KEY) === 'true') {
    dom.disclaimerModal.classList.add('hidden');
    return;
  }
  dom.disclaimerModal.classList.remove('hidden');
  dom.disclaimerAccept.addEventListener('click', () => {
    if (dom.disclaimerRemember.checked) {
      localStorage.setItem(DISCLAIMER_KEY, 'true');
    }
    dom.disclaimerModal.classList.add('hidden');
  });
}


/* ── 4b. TOOLTIPS ──────────────────────────────────────────── */

function wireTooltips() {
  const tooltipMap = {
    'imm-mode':    document.getElementById('tt-imm-mode'),
    'dental-method': document.getElementById('tt-dental-method'),
    'well-woman':  document.getElementById('tt-well-woman'),
  };

  document.querySelectorAll('.tooltip-trigger').forEach(trigger => {
    trigger.addEventListener('mouseenter', e => {
      const key     = trigger.dataset.tooltip;
      const content = tooltipMap[key];
      if (!content) return;
      dom.floatingTooltip.innerHTML  = content.innerHTML;
      dom.floatingTooltip.classList.add('visible');
      positionTooltip(e);
    });
    trigger.addEventListener('mousemove', positionTooltip);
    trigger.addEventListener('mouseleave', () => {
      dom.floatingTooltip.classList.remove('visible');
    });
  });

  function positionTooltip(e) {
    const tt  = dom.floatingTooltip;
    const pad = 14;
    let x = e.clientX + pad;
    let y = e.clientY + pad;
    // Keep within viewport
    if (x + tt.offsetWidth  > window.innerWidth  - 8) x = e.clientX - tt.offsetWidth  - pad;
    if (y + tt.offsetHeight > window.innerHeight - 8) y = e.clientY - tt.offsetHeight - pad;
    tt.style.left = x + 'px';
    tt.style.top  = y + 'px';
  }
}


/* ── 5. FILE UPLOAD ────────────────────────────────────────── */

function wireUploadHandlers() {
  dom.uploadArea.addEventListener('click', () => dom.fileInput.click());
  dom.fileInput.addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
  dom.uploadArea.addEventListener('dragover', e => { e.preventDefault(); dom.uploadArea.classList.add('drag-over'); });
  dom.uploadArea.addEventListener('dragleave', () => dom.uploadArea.classList.remove('drag-over'));
  dom.uploadArea.addEventListener('drop', e => {
    e.preventDefault(); dom.uploadArea.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
}

function handleFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    setUploadStatus('error','✗ Please upload an Excel file (.xlsx or .xls)');
    dom.uploadArea.classList.add('upload-error'); return;
  }
  setUploadStatus('','Reading file…');
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true,raw:false});
      if (!wb.SheetNames.includes(MRRS_SHEET_NAME)) {
        setUploadStatus('error',`✗ Sheet "${MRRS_SHEET_NAME}" not found.`);
        dom.uploadArea.classList.add('upload-error'); dom.uploadArea.classList.remove('upload-success'); return;
      }
      state.rawData = wb;
      const parsed  = parseMRRS(wb, false);
      if (!parsed.personnel.length) {
        setUploadStatus('error','✗ No personnel rows found.'); dom.uploadArea.classList.add('upload-error'); return;
      }
      state.personnel = parsed.personnel; state.colIndex = parsed.colIndex; state.parseStats = parsed.stats;
      dom.uploadArea.classList.add('upload-success'); dom.uploadArea.classList.remove('upload-error');
      setUploadStatus('success',
        `✓ ${file.name} — ${parsed.personnel.length} personnel loaded `+
        `(${parsed.stats.officers} officers, ${parsed.stats.enlisted} enlisted)`);
      dom.generateBtn.disabled = false;
      logParseResults(parsed);
    } catch(err) {
      console.error('File read error:',err);
      setUploadStatus('error','✗ Could not read file.'); dom.uploadArea.classList.add('upload-error');
    }
  };
  reader.onerror = () => setUploadStatus('error','✗ File read failed.');
  reader.readAsArrayBuffer(file);
}

function setUploadStatus(type,message) {
  dom.uploadStatus.textContent = message;
  dom.uploadStatus.className   = 'upload-status'+(type?` ${type}`:'');
}


/* ── 6. MRRS PARSER ────────────────────────────────────────── */

function parseMRRS(workbook, useMrrsDate) {
  const sheet = workbook.Sheets[MRRS_SHEET_NAME];
  const rows  = XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false});
  const headerRow = rows[MRRS_HEADER_ROW]||[];
  const colIndex  = {};
  headerRow.forEach((v,i)=>{ if(v!==null&&v!==undefined&&v!=='') colIndex[String(v).trim()]=i; });

  function getCell(row,key) {
    const idx=colIndex[COL[key]]; if(idx===undefined) return '';
    const v=row[idx]; return (v===null||v===undefined)?'':v;
  }

  function parseDate(val) {
    if(!val||val==='') return null;
    if(val instanceof Date) return isNaN(val.getTime())?null:val;
    if(typeof val==='number'){const d=XLSX.SSF.parse_date_code(val); return d?new Date(d.y,d.m-1,d.d):null;}
    if(typeof val==='string'){
      const t=val.trim(); if(!t) return null;
      const iso=Date.parse(t); if(!isNaN(iso)) return new Date(iso);
      const mdy=t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if(mdy) return new Date(+mdy[3],+mdy[1]-1,+mdy[2]);
      const dmy=t.match(/^(\d{1,2})-([A-Z]{3})-(\d{4})$/i);
      if(dmy){const m={JAN:0,FEB:1,MAR:2,APR:3,MAY:4,JUN:5,JUL:6,AUG:7,SEP:8,OCT:9,NOV:10,DEC:11}[dmy[2].toUpperCase()];
        if(m!==undefined) return new Date(+dmy[3],m,+dmy[1]);}
    }
    return null;
  }

  function dentalDueDate(row) {
    if(useMrrsDate) return parseDate(getCell(row,'DENTAL_DUE'));
    const dt=parseDate(getCell(row,'DENTAL_DT')); if(!dt) return null;
    const d=new Date(dt); d.setDate(d.getDate()+365); return d;
  }

  function parseDentalClass(val) {
    if(!val||String(val).trim()==='') return 4;
    const n=parseInt(String(val).trim(),10); return(n>=1&&n<=4)?n:4;
  }

  // MHA status text field. Returns { naFlag, due }
  //   ' Completed'        → OK        { naFlag:false, due:null, ok:true }
  //   ' Not Performed'    → OK        { naFlag:false, due:null, ok:true }
  //   '*Due DD-MMM-YYYY'  → evaluate  { naFlag:false, due:Date, ok:false }
  //   ' ' / blank         → NA        { naFlag:true,  due:null, ok:false }
  function parseMhaStatus(val) {
    const s = String(val||'').trim();
    if (!s) return { naFlag:true, due:null, ok:false };
    const lower = s.toLowerCase();
    if (lower === 'completed' || lower === 'not performed') return { naFlag:false, due:null, ok:true };
    // '*Due DD-MMM-YYYY'
    const m = s.match(/\*Due\s+(\d{1,2}-[A-Z]{3}-\d{4})/i);
    if (m) {
      const parts = m[1].split('-');
      const months = {JAN:0,FEB:1,MAR:2,APR:3,MAY:4,JUN:5,JUL:6,AUG:7,SEP:8,OCT:9,NOV:10,DEC:11};
      const month  = months[parts[1].toUpperCase()];
      const date   = (month !== undefined) ? new Date(+parts[2], month, +parts[0]) : null;
      return { naFlag:false, due:date, ok:false };
    }
    // Unrecognised value — treat as NA
    return { naFlag:true, due:null, ok:false };
  }

  function isYes(v){if(!v&&v!==0)return false;const s=String(v).trim().toLowerCase();return s==='yes'||s==='y'||s==='1'||s==='true';}
  function isDef(v){if(!v&&v!==0)return false;const s=String(v).trim().toLowerCase();return s==='yes'||s==='y'||s==='1'||s==='true'||s==='deferred';}

  function sect(row){
    const d=String(getCell(row,'COMP_DEPT')).trim(),p=String(getCell(row,'PLATOON')).trim();
    return(d&&p)?`${d}-${p}`:d||p||'';
  }

  const personnel=[]; let officers=0,enlisted=0,skipped=0;

  for(let i=MRRS_DATA_START;i<rows.length;i++){
    const row=rows[i];
    const name=String(getCell(row,'NAME')).trim();
    if(!name){skipped++;continue;}
    const offEnl=String(getCell(row,'OFF_ENL')).trim();
    const rank=String(getCell(row,'RANK')).trim();
    const isOfficer=offEnl.toLowerCase().includes('officer')||(offEnl===''&&rank.toUpperCase().includes('CWO'));
    if(isOfficer)officers++;else enlisted++;

    const bloodTypeRaw = String(getCell(row,'BLOOD_TYPE')).trim();

    personnel.push({
      name,rank,section:sect(row),
      offEnl:isOfficer?'Officer':'Enlisted',
      sex:String(getCell(row,'SEX')).trim(),
      imrStatus:String(getCell(row,'IMR_STATUS')).trim(),
      deploying:isYes(getCell(row,'DEPLOYING')),
      // Core
      phaDue:       parseDate(getCell(row,'PHA_DUE')),
      dentalDue:    dentalDueDate(row),
      dentalClass:  parseDentalClass(getCell(row,'DENTAL_CLASS')),
      hivDue:       parseDate(getCell(row,'HIV_DUE')),
      audioDue:     parseDate(getCell(row,'AUDIO_DUE')),
      // Deployment
      pdhaDue:      parseDate(getCell(row,'PDHA_DUE')),
      pdhraDue:     parseDate(getCell(row,'PDHRA_DUE')),
      mha2: parseMhaStatus(getCell(row,'MHA2_STATUS')),
      mha3: parseMhaStatus(getCell(row,'MHA3_STATUS')),
      mha4: parseMhaStatus(getCell(row,'MHA4_STATUS')),
      warningTagDue: parseDate(getCell(row,'WARNING_TAG_DUE')),
      verifyGlassesDue: parseDate(getCell(row,'VERIFY_GLASSES_DUE')),
      verifyInsertsDue: parseDate(getCell(row,'VERIFY_INSERTS_DUE')),
      // Accessions
      bloodType:    bloodTypeRaw,        // present = OK, blank = overdue
      refAudioDue:  parseDate(getCell(row,'REF_AUDIO_DUE')),
      dnaDue:       parseDate(getCell(row,'DNA_DUE')),
      g6pdDue:      parseDate(getCell(row,'G6PD_DUE')),
      sickleDue:    parseDate(getCell(row,'SICKLE_DUE')),
      // Other
      tstDue:       parseDate(getCell(row,'TST_DUE')),
      tstQuestDue:  parseDate(getCell(row,'TST_QUEST_DUE')),
      mammogramDue: parseDate(getCell(row,'MAMMOGRAM_DUE')),
      papDue:       parseDate(getCell(row,'PAP_DUE')),
      immunizations:{
        influenza:  {req:isYes(getCell(row,'INFLUENZA_REQ')),  due:parseDate(getCell(row,'INFLUENZA_DUE')),  deferred:isDef(getCell(row,'INFLUENZA_DEF')),  label:'Influenza'},
        tdap:       {req:isYes(getCell(row,'TDAP_REQ')),       due:parseDate(getCell(row,'TDAP_DUE')),       deferred:isDef(getCell(row,'TDAP_DEF')),       label:'TDap'},
        typhoid:    {req:isYes(getCell(row,'TYPHOID_REQ')),    due:parseDate(getCell(row,'TYPHOID_DUE')),    deferred:isDef(getCell(row,'TYPHOID_DEF')),    label:'Typhoid'},
        varicella:  {req:isYes(getCell(row,'VARICELLA_REQ')), due:parseDate(getCell(row,'VARICELLA_DUE')), deferred:isDef(getCell(row,'VARICELLA_DEF')), label:'Varicella'},
        mmr:        {req:isYes(getCell(row,'MMR_REQ')),        due:parseDate(getCell(row,'MMR_DUE')),        deferred:isDef(getCell(row,'MMR_DEF')),        label:'MMR'},
        hepa:       {req:isYes(getCell(row,'HEPA_REQ')),       due:parseDate(getCell(row,'HEPA_DUE')),       deferred:isDef(getCell(row,'HEPA_DEF')),       label:'Hep A'},
        hepb:       {req:isYes(getCell(row,'HEPB_REQ')),       due:parseDate(getCell(row,'HEPB_DUE')),       deferred:isDef(getCell(row,'HEPB_DEF')),       label:'Hep B'},
        twinrix:    {req:isYes(getCell(row,'TWINRIX_REQ')),    due:parseDate(getCell(row,'TWINRIX_DUE')),    deferred:isDef(getCell(row,'TWINRIX_DEF')),    label:'TwinRix'},
        rabies:     {req:isYes(getCell(row,'RABIES_REQ')),     due:parseDate(getCell(row,'RABIES_DUE')),     deferred:isDef(getCell(row,'RABIES_DEF')),     label:'Rabies'},
        rabiesTiter:{req:isYes(getCell(row,'RABIES_REQ')),     due:parseDate(getCell(row,'RABIES_TITER_DUE')),deferred:false,                               label:'Rabies Titer'},
        cholera:    {req:isYes(getCell(row,'CHOLERA_REQ')),    due:parseDate(getCell(row,'CHOLERA_DUE')),    deferred:isDef(getCell(row,'CHOLERA_DEF')),    label:'Cholera'},
        jev:        {req:isYes(getCell(row,'JEV_REQ')),        due:parseDate(getCell(row,'JEV_DUE')),        deferred:isDef(getCell(row,'JEV_DEF')),        label:'JEV'},
        mgc:        {req:isYes(getCell(row,'MGC_REQ')),        due:parseDate(getCell(row,'MGC_DUE')),        deferred:isDef(getCell(row,'MGC_DEF')),        label:'MGC'},
        polio:      {req:isYes(getCell(row,'POLIO_REQ')),      due:parseDate(getCell(row,'POLIO_DUE')),      deferred:isDef(getCell(row,'POLIO_DEF')),      label:'Polio'},
        yellowFever:{req:isYes(getCell(row,'YF_REQ')),         due:parseDate(getCell(row,'YF_DUE')),         deferred:isDef(getCell(row,'YF_DEF')),         label:'Yellow Fever'},
        anthrax:    {req:isYes(getCell(row,'ANTHRAX_REQ')),    due:parseDate(getCell(row,'ANTHRAX_DUE')),    deferred:isDef(getCell(row,'ANTHRAX_DEF')),    label:'Anthrax'},
        smallpox:   {req:isYes(getCell(row,'SMALLPOX_REQ')),   due:parseDate(getCell(row,'SMALLPOX_DUE')),   deferred:isDef(getCell(row,'SMALLPOX_DEF')),   label:'Smallpox'},
        adenovirus: {req:isYes(getCell(row,'ADENOVIRUS_REQ')), due:parseDate(getCell(row,'ADENOVIRUS_DUE')), deferred:isDef(getCell(row,'ADENOVIRUS_DEF')), label:'Adenovirus'},
        pneumo:     {req:isYes(getCell(row,'PNEUMO_REQ')),     due:parseDate(getCell(row,'PNEUMO_DUE')),     deferred:isDef(getCell(row,'PNEUMO_DEF')),     label:'Pneumococcal'},
        hpv:        {req:isYes(getCell(row,'HPV_REQ')),        due:parseDate(getCell(row,'HPV_DUE')),        deferred:isDef(getCell(row,'HPV_DEF')),        label:'HPV'},
      },
    });
  }

  return {personnel,colIndex,stats:{total:personnel.length,officers,enlisted,skipped,columns:Object.keys(colIndex).length}};
}

function logParseResults(parsed) {
  const {personnel,stats}=parsed;
  console.group('%cHitlist Hatcher — Parse Results','color:#1e3a5f;font-weight:bold;font-size:14px;');
  console.table({Total:stats.total,Officers:stats.officers,Enlisted:stats.enlisted,Skipped:stats.skipped});
  console.table(personnel.slice(0,5).map(p=>({name:p.name,rank:p.rank,dentalClass:p.dentalClass,bloodType:p.bloodType,mha2:p.mha2})));
  window._hhPersonnel=personnel;
  console.log('%cwindow._hhPersonnel available','color:#1e8449;font-style:italic;');
  console.groupEnd();
}


/* ── 7. READINESS LOGIC ────────────────────────────────────── */

function evaluateItem(dueDate, projDate, thresholds) {
  if(!dueDate||!(dueDate instanceof Date)||isNaN(dueDate)) return STATUS.NA;
  const days=daysDiff(dueDate,projDate);
  if(days<=0)                 return STATUS.OVERDUE;
  if(days<=thresholds.yellow) return STATUS.DUE_SOON;
  if(days<=thresholds.green)  return STATUS.UPCOMING;
  return STATUS.OK;
}

function evaluateDental(person, projDate, thresholds) {
  if(person.dentalClass===3) return {status:STATUS.OVERDUE, displayText:'Class 3'};
  if(person.dentalClass===4) return {status:STATUS.OVERDUE, displayText:'Class 4'};
  return {status:evaluateItem(person.dentalDue,projDate,thresholds), displayText:null};
}

// MHA: structured object from parseMhaStatus()
function evaluateMHA(mha, projDate, thresholds) {
  if (!mha || mha.naFlag)  return { status:STATUS.NA,      displayText:null };
  if (mha.ok)              return { status:STATUS.OK,      displayText:null };
  if (mha.due)             return { status:evaluateItem(mha.due, projDate, thresholds), displayText:null };
  // *Due present but date failed to parse
  return { status:STATUS.OVERDUE, displayText:'Due' };
}

// Blood Type: present = OK, blank = overdue. Show blood type value in cell.
function evaluateBloodType(bloodType) {
  if(bloodType && bloodType.trim()!=='') return {status:STATUS.OK, displayText:bloodType.trim()};
  return {status:STATUS.OVERDUE, displayText:'Missing'};
}

// Accession date items: blank = NA (not required). Has date = evaluate normally.
function evaluateAccessionDate(dueDate, projDate, thresholds) {
  if (!dueDate || !(dueDate instanceof Date) || isNaN(dueDate)) return STATUS.NA;
  return evaluateItem(dueDate, projDate, thresholds);
}

function evaluateImmunization(immObj,projDate,thresholds){
  if(!immObj.req)     return STATUS.NA;
  if(immObj.deferred) return STATUS.NA;
  if(!immObj.due)     return STATUS.OVERDUE;
  return evaluateItem(immObj.due,projDate,thresholds);
}

function evaluateImmunizations(person,selectedKeys,projDate,thresholds){
  let worstStatus=STATUS.NA,count=0;
  const dueNames=[],perVaccine={};
  selectedKeys.forEach(key=>{
    const imm=person.immunizations[key]; if(!imm) return;
    const status=evaluateImmunization(imm,projDate,thresholds);
    perVaccine[key]={status,due:imm.due};
    if(SEVERITY[status]>SEVERITY[worstStatus]) worstStatus=status;
    if(status===STATUS.OVERDUE||status===STATUS.DUE_SOON||status===STATUS.UPCOMING){count++;dueNames.push(imm.label);}
  });
  return {worstStatus,count,dueNames,perVaccine};
}

function evaluatePerson(person, settings, projDate) {
  const t=settings.thresholds||DEFAULT_THRESHOLDS;
  const ei=due=>({status:evaluateItem(due,projDate,t),displayText:null});
  const na={status:STATUS.NA,displayText:null};

  const phaResult        = settings.items.pha          ? ei(person.phaDue)           : na;
  const dentalResult     = settings.items.dental        ? evaluateDental(person,projDate,t) : na;
  const hivResult        = settings.items.hiv           ? ei(person.hivDue)           : na;
  const audioResult      = settings.items.audio         ? ei(person.audioDue)         : na;
  const pdhaResult       = settings.items.pdha          ? ei(person.pdhaDue)          : na;
  const pdhraResult      = settings.items.pdhra         ? ei(person.pdhraDue)         : na;
  const mha2Result       = settings.items.mha2          ? evaluateMHA(person.mha2, projDate, t) : na;
  const mha3Result       = settings.items.mha3          ? evaluateMHA(person.mha3, projDate, t) : na;
  const mha4Result       = settings.items.mha4          ? evaluateMHA(person.mha4, projDate, t) : na;
  const warningTagResult = settings.items.warningTag    ? ei(person.warningTagDue)    : na;
  const verifyGlassesResult = settings.items.verifyGlasses ? ei(person.verifyGlassesDue) : na;
  const verifyInsertsResult = settings.items.verifyInserts  ? ei(person.verifyInsertsDue)  : na;

  const bloodTypeResult  = settings.items.bloodType     ? evaluateBloodType(person.bloodType) : na;
  const refAudioResult   = settings.items.refAudio      ? {status:evaluateAccessionDate(person.refAudioDue,projDate,t),displayText:null} : na;
  const dnaResult        = settings.items.dna           ? {status:evaluateAccessionDate(person.dnaDue,projDate,t),displayText:null} : na;
  const g6pdResult       = settings.items.g6pd          ? {status:evaluateAccessionDate(person.g6pdDue,projDate,t),displayText:null} : na;
  const sickleResult     = settings.items.sickle        ? {status:evaluateAccessionDate(person.sickleDue,projDate,t),displayText:null} : na;

  const tstResult        = settings.items.tst           ? ei(person.tstDue)           : na;
  const tstQuestResult   = settings.items.tstQuest      ? ei(person.tstQuestDue)      : na;

  let wellWomanResult = na;
  if(settings.items.wellWoman && person.sex.toLowerCase()==='female'){
    const ms=evaluateItem(person.mammogramDue,projDate,t), ps=evaluateItem(person.papDue,projDate,t);
    wellWomanResult={status:SEVERITY[ms]>=SEVERITY[ps]?ms:ps, displayText:null};
  }

  const immKeys=settings.immunizationKeys||[];
  const immRes=evaluateImmunizations(person,immKeys,projDate,t);
  const immunizationsResult={
    grouped:{status:immRes.worstStatus,count:immRes.count,dueNames:immRes.dueNames},
    perVaccine:immRes.perVaccine,
  };

  const allStatuses=[
    phaResult.status,dentalResult.status,hivResult.status,audioResult.status,
    pdhaResult.status,pdhraResult.status,
    mha2Result.status,mha3Result.status,mha4Result.status,
    warningTagResult.status,verifyGlassesResult.status,verifyInsertsResult.status,
    bloodTypeResult.status,refAudioResult.status,dnaResult.status,g6pdResult.status,sickleResult.status,
    tstResult.status,tstQuestResult.status,wellWomanResult.status,immRes.worstStatus,
  ];
  const overallStatus=allStatuses.reduce((w,s)=>SEVERITY[s]>SEVERITY[w]?s:w,STATUS.NA);
  const isDue=SEVERITY[overallStatus]>=SEVERITY[STATUS.UPCOMING];

  return {person,overallStatus,isDue,items:{
    pha:phaResult,dental:dentalResult,hiv:hivResult,audio:audioResult,
    pdha:pdhaResult,pdhra:pdhraResult,
    mha2:mha2Result,mha3:mha3Result,mha4:mha4Result,
    warningTag:warningTagResult,verifyGlasses:verifyGlassesResult,verifyInserts:verifyInsertsResult,
    bloodType:bloodTypeResult,refAudio:refAudioResult,dna:dnaResult,g6pd:g6pdResult,sickle:sickleResult,
    tst:tstResult,tstQuest:tstQuestResult,wellWoman:wellWomanResult,
    immunizations:immunizationsResult,
  }};
}

function applyFilters(personnel,settings,projDate){
  const allEval=personnel.map(p=>evaluatePerson(p,settings,projDate));
  const stats=computeReadinessStats(personnel);
  const filter=settings.offEnlFilter||'combined';
  let filtered=allEval.filter(r=>{
    if(!r.isDue) return false;
    if(filter==='officer')  return r.person.offEnl==='Officer';
    if(filter==='enlisted') return r.person.offEnl==='Enlisted';
    return true;
  });
  const sortBy=settings.sortBy||'name';
  filtered.sort((a,b)=>{
    if(sortBy==='section'){const c=a.person.section.localeCompare(b.person.section);if(c!==0)return c;}
    return a.person.name.localeCompare(b.person.name);
  });
  return {results:filtered,stats};
}

function computeReadinessStats(personnel){
  let fullyReady=0,notReady=0,partial=0,indeterminate=0;
  personnel.forEach(p=>{
    const s=(p.imrStatus||'').toLowerCase();
    if(s.includes('fully'))fullyReady++;
    else if(s.includes('not'))notReady++;
    else if(s.includes('partial'))partial++;
    else indeterminate++;
  });
  const total=personnel.length,pct=n=>total>0?((n/total)*100).toFixed(1):'0.0';
  return{total,fullyReady,notReady,partial,indeterminate,
    fullyReadyPct:pct(fullyReady),notReadyPct:pct(notReady),
    partialPct:pct(partial),indeterminatePct:pct(indeterminate)};
}


/* ── 8. REPORT RENDERER ────────────────────────────────────── */

function renderReport(results,stats,settings,projDate){
  if(results.length===0){
    dom.previewPlaceholder.classList.add('hidden');
    dom.reportOutput.innerHTML=`
      <div style="background:#fff;border-radius:6px;padding:48px;text-align:center;box-shadow:0 2px 10px rgba(0,0,0,.1);">
        <div style="font-size:48px;margin-bottom:16px;">✓</div>
        <p style="font-size:18px;font-weight:700;color:#1e3a5f;margin-bottom:8px;">No personnel due</p>
        <p style="font-size:14px;color:#444;">No one is due or upcoming for the selected items within the current warning windows.</p>
      </div>`;
    dom.exportBar.hidden=true; return;
  }

  if(settings.offEnlFilter==='separate'){
    const officers=results.filter(r=>r.person.offEnl==='Officer');
    const enlisted=results.filter(r=>r.person.offEnl==='Enlisted');
    const colDefs=getColumnDefs(settings); updateColumnCounter(colDefs.length);
    let html='';
    if(officers.length) html+=buildHitListHTML(officers,stats,settings,projDate,colDefs,'Officer Hit List');
    if(enlisted.length) html+=`<div style="margin-top:32px;">${buildHitListHTML(enlisted,stats,settings,projDate,colDefs,'Enlisted Hit List')}</div>`;
    dom.previewPlaceholder.classList.add('hidden');
    dom.reportOutput.innerHTML=html;
    dom.exportBar.hidden=false; state.reportGenerated=true; return;
  }

  const colDefs=getColumnDefs(settings); updateColumnCounter(colDefs.length);
  dom.previewPlaceholder.classList.add('hidden');
  dom.reportOutput.innerHTML=buildHitListHTML(results,stats,settings,projDate,colDefs,null);
  dom.exportBar.hidden=false; state.reportGenerated=true;
}

function buildHitListHTML(results,stats,settings,projDate,colDefs,subtitleOverride){
  return`<div class="hitlist-wrapper">
    ${buildHeader(stats,settings,projDate,subtitleOverride)}
    ${buildLegend(settings)}
    <div class="hitlist-table-wrapper">${buildTable(results,colDefs,settings,projDate)}</div>
  </div>`;
}

function getColumnDefs(settings){
  const d=[];
  d.push({key:'name',label:'Name',type:'identity'});
  if(settings.showRank)    d.push({key:'rank',   label:'Rank',   type:'identity'});
  if(settings.showSection) d.push({key:'section',label:'Section',type:'identity'});
  // Core
  if(settings.items.pha)          d.push({key:'pha',          label:'PHA',         type:'item'});
  if(settings.items.dental)       d.push({key:'dental',       label:'Dental',      type:'item'});
  if(settings.items.hiv)          d.push({key:'hiv',          label:'HIV Lab',     type:'item'});
  if(settings.items.audio)        d.push({key:'audio',        label:'Audiogram',   type:'item'});
  // Deployment
  if(settings.items.pdha)         d.push({key:'pdha',         label:'PDHA',        type:'item'});
  if(settings.items.pdhra)        d.push({key:'pdhra',        label:'PDHRA',       type:'item'});
  if(settings.items.mha2)         d.push({key:'mha2',         label:'MHA 2',       type:'item'});
  if(settings.items.mha3)         d.push({key:'mha3',         label:'MHA 3',       type:'item'});
  if(settings.items.mha4)         d.push({key:'mha4',         label:'MHA 4',       type:'item'});
  if(settings.items.verifyGlasses)d.push({key:'verifyGlasses',label:'Glasses',     type:'item'});
  if(settings.items.verifyInserts)d.push({key:'verifyInserts',label:'Inserts',     type:'item'});
  if(settings.items.warningTag)   d.push({key:'warningTag',   label:'Warn Tag',    type:'item'});
  // Accessions
  if(settings.items.bloodType)    d.push({key:'bloodType',    label:'Blood Type',  type:'item'});
  if(settings.items.refAudio)     d.push({key:'refAudio',     label:'Ref Audio',   type:'item'});
  if(settings.items.dna)          d.push({key:'dna',          label:'DNA',         type:'item'});
  if(settings.items.g6pd)         d.push({key:'g6pd',         label:'G6PD',        type:'item'});
  if(settings.items.sickle)       d.push({key:'sickle',       label:'Sickle Cell', type:'item'});
  // Other
  if(settings.items.tst)          d.push({key:'tst',          label:'TST',         type:'item'});
  if(settings.items.tstQuest)     d.push({key:'tstQuest',     label:'TST Quest',   type:'item'});
  if(settings.items.wellWoman)    d.push({key:'wellWoman',    label:'Well-Woman',  type:'item'});
  // Immunizations
  if(settings.items.immunizations){
    if((settings.immunDisplayMode||'grouped')==='grouped'){
      d.push({key:'imm-grouped',label:'Immunizations',type:'imm-grouped'});
    } else {
      (settings.immunizationKeys||[]).forEach(k=>{
        d.push({key:`imm-${k}`,label:IMMUNIZATION_LABELS[k]||k,type:'imm-individual',vaccKey:k});
      });
    }
  }
  return d;
}

function buildHeader(stats,settings,projDate,subtitleOverride){
  const unitName  =escHtml(settings.unitName||'Unit Name');
  const projStr   =formatDateFull(projDate);
  const reportDate=formatDateFull(new Date());

  const emblem=settings.emblemBase64
    ?`<img class="hitlist-emblem" src="${settings.emblemBase64}" alt="Unit Emblem" />`
    :`<div class="hitlist-emblem-placeholder">Unit<br/>Emblem</div>`;

  const titleLine=subtitleOverride
    ?`<div class="hitlist-title-line">${escHtml(subtitleOverride)}</div>`
    :`<div class="hitlist-title-line">Report Date: ${reportDate}</div>`;

  // Multi-column info bar
  const infoBar = buildInfoBar(settings);

  return`<div class="hitlist-header">
    <div class="hitlist-header-top">${emblem}
      <div class="hitlist-title-block">
        <div class="hitlist-unit-name">${unitName} Medical Hit List</div>
        ${titleLine}
        <div class="hitlist-projection-line">Projected to: ${projStr}</div>
      </div>${emblem}
    </div>
    <div class="hitlist-stats-bar">
      <div class="hitlist-stat"><span class="hitlist-stat-label">Fully Ready</span><span class="hitlist-stat-value stat-good">${stats.fullyReadyPct}%</span></div>
      <div class="hitlist-stat"><span class="hitlist-stat-label">Not Ready</span><span class="hitlist-stat-value stat-bad">${stats.notReadyPct}%</span></div>
      <div class="hitlist-stat"><span class="hitlist-stat-label">Partial</span><span class="hitlist-stat-value stat-warn">${stats.partialPct}%</span></div>
      <div class="hitlist-stat"><span class="hitlist-stat-label">Indeterminate</span><span class="hitlist-stat-value stat-neutral">${stats.indeterminatePct}%</span></div>
    </div>
    ${infoBar}
  </div>`;
}

function buildInfoBar(settings){
  const cols=settings.infoColumns||[];
  if(!cols.length||cols.every(c=>!c.text.trim())) return '';
  const visibleCols=cols.filter(c=>c.text.trim());
  if(!visibleCols.length) return '';
  const cells=visibleCols.map(col=>{
    const alignClass=col.align==='center'?'info-col-center':col.align==='right'?'info-col-right':'info-col-left';
    return`<div class="info-col ${alignClass}">${escHtml(col.text).replace(/\n/g,'<br/>')}</div>`;
  }).join('');
  return`<div class="hitlist-info-bar hitlist-info-grid">${cells}</div>`;
}

function buildLegend(settings){
  const t=settings.thresholds||DEFAULT_THRESHOLDS;
  return`<div class="hitlist-legend">
    <span class="hitlist-legend-label">Key:</span>
    <span class="legend-item"><span class="legend-swatch red"></span>Overdue / Class 3–4</span>
    <span class="legend-item"><span class="legend-swatch yellow"></span>Due within ${t.yellow} days</span>
    <span class="legend-item"><span class="legend-swatch green"></span>Due within ${t.green} days</span>
    <span class="legend-item"><span class="legend-swatch none" style="border:1px solid #c8ceda;"></span>Not due</span>
  </div>`;
}

function buildTable(results,colDefs,settings,projDate){
  const hdr=colDefs.map(col=>`<th class="${col.type==='identity'&&col.key==='name'?'col-name-hdr':''}">${escHtml(col.label)}</th>`).join('');
  const rows=results.map(r=>`<tr>${colDefs.map(col=>getCellHTML(r,col,settings,projDate)).join('')}</tr>`).join('');
  return`<table class="hitlist-table"><thead><tr>${hdr}</tr></thead><tbody>${rows}</tbody></table>`;
}

function getCellHTML(result,col,settings,projDate){
  const{key,type,vaccKey}=col;
  if(type==='identity'){
    if(key==='name')    return`<td class="col-name">${escHtml(result.person.name)}</td>`;
    if(key==='rank')    return`<td class="col-rank">${escHtml(result.person.rank)}</td>`;
    if(key==='section') return`<td class="col-section">${escHtml(result.person.section)}</td>`;
  }
  if(type==='imm-grouped'){
    const grp=result.items.immunizations.grouped;
    if(grp.status===STATUS.NA||grp.count===0) return`<td class="col-item cell-na">—</td>`;
    const tooltip=grp.dueNames.map(n=>escHtml(n)).join('<br/>');
    return`<td class="col-item ${statusToCss(grp.status)}"><div class="imm-grouped">${grp.count} due<div class="imm-tooltip">${tooltip}</div></div></td>`;
  }
  if(type==='imm-individual'){
    const pv=result.items.immunizations.perVaccine[vaccKey];
    if(!pv||pv.status===STATUS.NA||pv.status===STATUS.OK) return`<td class="col-item cell-na">—</td>`;
    return`<td class="col-item ${statusToCss(pv.status)}">${pv.due?formatDate(pv.due):'?'}</td>`;
  }
  if(type==='item'){
    const ir=result.items[key];
    if(!ir||ir.status===STATUS.NA||ir.status===STATUS.OK) return`<td class="col-item cell-na">—</td>`;
    // Items with displayText override (dental class, MHA, blood type)
    if(ir.displayText) return`<td class="col-item ${statusToCss(ir.status)}">${escHtml(ir.displayText)}</td>`;
    // Items with no due date (accessions: blank = overdue with no date to show)
    const due=getDueDateForItem(result.person,key);
    const dateStr=due?formatDate(due):'Due';
    return`<td class="col-item ${statusToCss(ir.status)}">${dateStr}</td>`;
  }
  return`<td class="col-item cell-na">—</td>`;
}

function getDueDateForItem(person,key){
  const map={
    pha:person.phaDue, dental:person.dentalDue, hiv:person.hivDue, audio:person.audioDue,
    pdha:person.pdhaDue, pdhra:person.pdhraDue,
    mha2:person.mha2?.due||null, mha3:person.mha3?.due||null, mha4:person.mha4?.due||null,
    warningTag:person.warningTagDue, verifyGlasses:person.verifyGlassesDue, verifyInserts:person.verifyInsertsDue,
    refAudio:person.refAudioDue, dna:person.dnaDue, g6pd:person.g6pdDue, sickle:person.sickleDue,
    tst:person.tstDue, tstQuest:person.tstQuestDue,
    wellWoman:person.mammogramDue||person.papDue,
  };
  return map[key]||null;
}

function statusToCss(status){
  if(status===STATUS.OVERDUE)  return'status-red';
  if(status===STATUS.DUE_SOON) return'status-yellow';
  if(status===STATUS.UPCOMING) return'status-green';
  return'';
}

function escHtml(str){
  if(!str) return'';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}


/* ── 9. SETTINGS PANEL ─────────────────────────────────────── */

function initSettingsPanel(){

  // Build immunization checkboxes
  const frag=document.createDocumentFragment();
  IMMUNIZATION_KEYS.forEach(key=>{
    const label=document.createElement('label');
    label.className='cb-label cb-label-imm';
    label.innerHTML=`<input type="checkbox" class="imm-cb" data-key="${key}"${IMMUNIZATION_DEFAULTS[key]?' checked':''}/>
      ${escHtml(IMMUNIZATION_LABELS[key])}`;
    frag.appendChild(label);
  });
  dom.immCheckboxes.appendChild(frag);

  // Dental sub-toggle
  dom.itemDental.addEventListener('change',()=>{
    dom.dentalSub.classList.toggle('hidden',!dom.itemDental.checked);
    onSettingsChanged();
  });

  // Immunization mode toggle
  dom.immModeGroupedBtn.addEventListener('click',()=>{
    state.immDisplayMode='grouped';
    dom.immModeGroupedBtn.classList.add('active'); dom.immModeIndividualBtn.classList.remove('active');
    refreshColumnCounter();
  });
  dom.immModeIndividualBtn.addEventListener('click',()=>{
    state.immDisplayMode='individual';
    dom.immModeIndividualBtn.classList.add('active'); dom.immModeGroupedBtn.classList.remove('active');
    refreshColumnCounter();
  });

  // Select / Clear all immunizations
  dom.immSelectAll.addEventListener('click',()=>{dom.immCheckboxes.querySelectorAll('.imm-cb').forEach(cb=>cb.checked=true);refreshColumnCounter();});
  dom.immClearAll.addEventListener('click', ()=>{dom.immCheckboxes.querySelectorAll('.imm-cb').forEach(cb=>cb.checked=false);refreshColumnCounter();});

  // All checkboxes → column counter
  document.querySelectorAll('input[type="checkbox"]').forEach(cb=>cb.addEventListener('change',onSettingsChanged));

  // Selects
  dom.offEnlFilter.addEventListener('change',onSettingsChanged);
  dom.sortBy.addEventListener('change',onSettingsChanged);

  // Threshold validation
  const validateThresholds=debounce(()=>{
    const y=parseInt(dom.threshYellow.value,10),g=parseInt(dom.threshGreen.value,10);
    if(isNaN(y)||isNaN(g)||y<1||g<1) dom.threshError.textContent='Thresholds must be positive numbers.';
    else if(y>=g) dom.threshError.textContent='Yellow window must be smaller than Green window.';
    else dom.threshError.textContent='';
    refreshColumnCounter();
  },DEBOUNCE_MS);
  dom.threshYellow.addEventListener('input',validateThresholds);
  dom.threshGreen.addEventListener('input',validateThresholds);

  // Emblem upload
  dom.emblemInput.addEventListener('change',e=>{
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=ev=>{
      state.emblemBase64=ev.target.result;
      dom.emblemPreview.style.backgroundImage=`url('${ev.target.result}')`;
      dom.emblemPreview.classList.add('has-image'); dom.clearEmblemBtn.classList.remove('hidden');
    };
    reader.readAsDataURL(file);
  });
  dom.clearEmblemBtn.addEventListener('click',()=>{
    state.emblemBase64=null; dom.emblemPreview.style.backgroundImage='';
    dom.emblemPreview.classList.remove('has-image'); dom.clearEmblemBtn.classList.add('hidden');
    dom.emblemInput.value='';
  });

  // Dental date radio
  document.querySelectorAll('input[name="dentalDate"]').forEach(r=>r.addEventListener('change',onSettingsChanged));

  // Text inputs
  dom.unitName.addEventListener('input',debounce(onSettingsChanged,DEBOUNCE_MS));

  // Info column builder
  initInfoColumns();
  dom.infoColCount.addEventListener('change',initInfoColumns);

  refreshColumnCounter();
}

/*
  Info text multi-column editor.
  Builds N column inputs (textarea + alignment select) based on the
  selected column count.
*/
function initInfoColumns(){
  const count=parseInt(dom.infoColCount.value,10)||1;
  // Preserve existing text values before rebuilding
  const existing=[];
  dom.infoColContainer.querySelectorAll('.info-col-editor').forEach(ed=>{
    existing.push({
      text: ed.querySelector('textarea').value,
      align: ed.querySelector('select').value,
    });
  });
  dom.infoColContainer.innerHTML='';
  for(let i=0;i<count;i++){
    const prev=existing[i]||{text:'',align:'left'};
    const wrapper=document.createElement('div');
    wrapper.className='info-col-editor';
    wrapper.innerHTML=`
      <div class="info-col-header">
        <span class="info-col-num">Column ${i+1}</span>
        <select class="info-align-select">
          <option value="left"${prev.align==='left'?' selected':''}>Left</option>
          <option value="center"${prev.align==='center'?' selected':''}>Center</option>
          <option value="right"${prev.align==='right'?' selected':''}>Right</option>
        </select>
      </div>
      <textarea class="settings-textarea info-col-textarea" rows="4"
        placeholder="Clinic hours, phone numbers...">${escHtml(prev.text)}</textarea>`;
    dom.infoColContainer.appendChild(wrapper);
  }
}

function getInfoColumns(){
  const cols=[];
  dom.infoColContainer.querySelectorAll('.info-col-editor').forEach(ed=>{
    cols.push({
      text:  ed.querySelector('textarea').value,
      align: ed.querySelector('select').value,
    });
  });
  return cols;
}

function onSettingsChanged(){ refreshColumnCounter(); }

function getSettingsFromUI(){
  const immunizationKeys=[];
  dom.immCheckboxes.querySelectorAll('.imm-cb:checked').forEach(cb=>immunizationKeys.push(cb.dataset.key));

  const y=parseInt(dom.threshYellow.value,10), g=parseInt(dom.threshGreen.value,10);
  const thresholds=(!isNaN(y)&&!isNaN(g)&&y>=1&&g>y)?{yellow:y,green:g}:DEFAULT_THRESHOLDS;

  const dentalRadio=document.querySelector('input[name="dentalDate"]:checked');
  const dentalUseMrrsDate=dentalRadio?dentalRadio.value==='mrrs':false;

  let projectionDate=new Date();
  if(dom.projectionDate.value){
    const pd=new Date(dom.projectionDate.value+'T00:00:00');
    if(!isNaN(pd)) projectionDate=pd;
  }

  return{
    unitName:    dom.unitName.value.trim(),
    infoColumns: getInfoColumns(),
    emblemBase64: state.emblemBase64,
    items:{
      pha:           dom.itemPha.checked,
      dental:        dom.itemDental.checked,
      hiv:           dom.itemHiv.checked,
      audio:         dom.itemAudio.checked,
      pdha:          dom.itemPdha.checked,
      pdhra:         dom.itemPdhra.checked,
      mha2:          dom.itemMha2.checked,
      mha3:          dom.itemMha3.checked,
      mha4:          dom.itemMha4.checked,
      warningTag:    dom.itemWarningTag.checked,
      verifyGlasses: dom.itemVerifyGlasses.checked,
      verifyInserts: dom.itemVerifyInserts.checked,
      bloodType:     dom.itemBloodType.checked,
      refAudio:      dom.itemRefAudio.checked,
      dna:           dom.itemDna.checked,
      g6pd:          dom.itemG6pd.checked,
      sickle:        dom.itemSickle.checked,
      tst:           dom.itemTst.checked,
      tstQuest:      dom.itemTstQuest.checked,
      wellWoman:     dom.itemWellWoman.checked,
      immunizations: immunizationKeys.length>0,
    },
    immunizationKeys, immunDisplayMode:state.immDisplayMode,
    thresholds, offEnlFilter:dom.offEnlFilter.value, sortBy:dom.sortBy.value,
    showRank:dom.showRank.checked, showSection:dom.showSection.checked,
    dentalUseMrrsDate, projectionDate,
  };
}

function refreshColumnCounter(){
  const settings=getSettingsFromUI();
  updateColumnCounter(getColumnDefs(settings).length);
}


/* ── Generate button ──────────────────────────────────────── */

dom.generateBtn.addEventListener('click',()=>{
  if(!state.rawData||!state.personnel.length){alert('Please upload a MRRS file first.');return;}

  const settings=getSettingsFromUI();

  // Re-parse if dental date method changed
  if(settings.dentalUseMrrsDate!==state._lastDentalFlag){
    const reparsed=parseMRRS(state.rawData, settings.dentalUseMrrsDate);
    state.personnel=reparsed.personnel; state._lastDentalFlag=settings.dentalUseMrrsDate;
  }

  const{results,stats}=applyFilters(state.personnel,settings,settings.projectionDate);
  state.filteredResults=results; state.currentStats=stats;
  renderReport(results,stats,settings,settings.projectionDate);
});


/* ── 10. SETTINGS PERSISTENCE ──────────────────────────────── */

function wireSettingsHandlers(){
  dom.exportSettingsBtn.addEventListener('click',()=>alert('Settings export will be available in Phase 5.'));
  dom.importSettingsInput.addEventListener('change',()=>alert('Settings import will be available in Phase 5.'));
}


/* ── 11. PDF EXPORT ────────────────────────────────────────── */

function wireExportPdfHandler(){
  dom.exportPdfBtn.addEventListener('click',()=>{
    dom.printOutput.innerHTML=dom.reportOutput.innerHTML;
    window.print();
  });
}


/* ── 12. UTILITIES ─────────────────────────────────────────── */

function setDefaultProjectionDate(){
  const today=new Date();
  const y=today.getMonth()===11?today.getFullYear()+1:today.getFullYear();
  const m=(today.getMonth()+1)%12;
  const lastDay=new Date(y,m+1,0).getDate();
  dom.projectionDate.value=`${y}-${String(m+1).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;
}

function formatBytes(b){return b<1024?`${b} B`:b<1048576?`${(b/1024).toFixed(1)} KB`:`${(b/1048576).toFixed(1)} MB`;}

function formatDate(date){
  if(!date||!(date instanceof Date)||isNaN(date))return'';
  const m=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return`${date.getDate()} ${m[date.getMonth()]}`;
}

function formatDateFull(date){
  if(!date||!(date instanceof Date)||isNaN(date))return'';
  const m=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return`${date.getDate()} ${m[date.getMonth()]} ${date.getFullYear()}`;
}

function daysDiff(a,b){return Math.round((a-b)/86400000);}

function debounce(fn,delay){let t;return(...args)=>{clearTimeout(t);t=setTimeout(()=>fn(...args),delay);};}

function updateColumnCounter(count){
  dom.columnCount.textContent=count;
  dom.columnCounter.classList.remove('warn-yellow','warn-red');
  if(count>=COL_WARN_RED)    dom.columnCounter.classList.add('warn-red');
  else if(count>=COL_WARN_YELLOW) dom.columnCounter.classList.add('warn-yellow');
}
