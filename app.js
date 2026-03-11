/* ============================================================
   HITLIST HATCHER 3.0 — APPLICATION LOGIC

   Copyright © 2026 Hitlist Hatcher. All rights reserved.
   This software is protected under copyright law. Unauthorized
   copying, modification, or distribution is strictly prohibited
   without the express written permission of the author.
   This work was created independently, outside the scope of any
   official military duty, and is not a work of the U.S. Government.
============================================================ */

/* ── 1. CONSTANTS ──────────────────────────────────────────── */

const APP_VERSION = '3.0.0';
const MRRS_HEADER_ROW = 2;
const MRRS_DATA_START = 3;
const MRRS_SHEET_NAME = 'IMR Detail';

// ── MRRS format validation ──
// Rows 0–1 are expected to contain FOUO / Privacy Act warning text.
// Row 2 (MRRS_HEADER_ROW) contains column headers.
// Update these if MRRS changes their export format.
const MRRS_FOUO_KEYWORDS     = ['FOUO', 'Privacy Act'];
const MRRS_SIGNATURE_COLUMNS = ['Name', 'Rank/Rate', 'IMR Status', 'PHA Due', 'Dental Exam Due', 'Off Enl Indicator'];
const MRRS_FORMAT_ERROR       = '✗ This file doesn\'t match the expected MRRS format. Hitlist Hatcher requires the Excel IMR Report (Details) report.\n\nIn MRRS: Reports → IMR → Force → IMR, under Report Title select Excel IMR Report (Details), within Unit select your command, then click Run.';

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
  INFLUENZA_DUE:'Influenza Due',
  TDAP_DUE:'Tet/Dipth Due',
  TYPHOID_DUE:'Typhoid Due',
  VARICELLA_DUE:'Varicella Due',
  MMR_DUE:'MMR Due',
  HEPA_DUE:'HepA Due',
  HEPB_DUE:'HepB Due',
  TWINRIX_DUE:'TwinRix Due',
  RABIES_DUE:'Rabies Due',
  RABIES_TITER_DUE:'Rabies Titer Due',
  CHOLERA_DUE:'Cholera Due',
  JEV_DUE:'JEV Due',
  MGC_DUE:'MGC Due',
  POLIO_DUE:'Polio Due',
  YF_DUE:'Yellow Fever Due',
  ANTHRAX_DUE:'Anthrax Due',
  SMALLPOX_DUE:'Smallpox Due',
  ADENOVIRUS_DUE:'Adenovirus Due',
  PNEUMO_DUE:'Pneumococcal Due',
  HPV_DUE:'HPV Due',
};

const STATUS   = { NA:'NA', OK:'OK', UPCOMING:'UPCOMING', DUE_SOON:'DUE_SOON', OVERDUE:'OVERDUE' };
const SEVERITY = { NA:0, OK:1, UPCOMING:2, DUE_SOON:3, OVERDUE:4 };
const DEFAULT_THRESHOLDS = { yellow:7, green:30 };

const STORAGE_KEY      = 'hitlistHatcher_settings_v3';
const DISCLAIMER_KEY   = 'hitlistHatcher_disclaimerSeen';
const WALKTHROUGH_KEY  = 'hitlistHatcher_walkthroughSeen';
const COL_WARN_YELLOW  = 12;
const COL_WARN_RED     = 14;
const EMBLEM_MAX_BYTES = 153600;  // 150KB raw file size limit
const DEBOUNCE_MS      = 400;

const IMMUNIZATION_KEYS = [
  'adenovirus','anthrax','cholera','hepa','hepb','hpv',
  'influenza','jev','mgc','mmr','pneumo','polio',
  'rabies','rabiesTiter','smallpox','tdap','twinrix',
  'typhoid','varicella','yellowFever',
];

const IMMUNIZATION_DEFAULTS = {
  adenovirus:true, anthrax:true, cholera:true, hepa:true, hepb:true, hpv:true,
  influenza:true,  jev:true,     mgc:true,     mmr:true,  pneumo:true, polio:true,
  rabies:true,     rabiesTiter:true, smallpox:true, tdap:true, twinrix:true,
  typhoid:true,    varicella:true,   yellowFever:true,
};

const IMMUNIZATION_LABELS = {
  adenovirus:'Adenovirus', anthrax:'Anthrax',     cholera:'Cholera',
  hepa:'Hep A',            hepb:'Hep B',          hpv:'HPV',
  influenza:'Influenza',   jev:'JEV',             mgc:'MGC',
  mmr:'MMR',               pneumo:'Pneumococcal', polio:'Polio',
  rabies:'Rabies',         rabiesTiter:'Rabies Titer', smallpox:'Smallpox',
  tdap:'TDap',             twinrix:'TwinRix',     typhoid:'Typhoid',
  varicella:'Varicella',   yellowFever:'Yellow Fever',
};


/* ── 1b. ERROR CAPTURE ─────────────────────────────────────── */

const errorLog = [];
const MAX_ERROR_LOG = 10;

// Add an error entry. Deduplicates by type+message — if the last entry
// matches, its count is incremented instead of adding a duplicate.
function captureError(entry) {
  entry.timestamp = new Date().toISOString();
  entry.appVersion = APP_VERSION;
  const last = errorLog.length ? errorLog[errorLog.length - 1] : null;
  if (last && last.type === entry.type && last.message === entry.message) {
    last.count = (last.count || 1) + 1;
    last.timestamp = entry.timestamp; // update to latest occurrence
    return;
  }
  entry.count = 1;
  if (errorLog.length >= MAX_ERROR_LOG) errorLog.shift();
  errorLog.push(entry);
}

// Build a diagnostics snapshot for feedback submissions.
// No PII — only structural/environment info.
function getErrorReport() {
  return {
    appVersion: APP_VERSION,
    browser: navigator.userAgent,
    screen: `${screen.width}x${screen.height}`,
    viewport: `${window.innerWidth}x${window.innerHeight}`,
    personnelLoaded: state.personnel.length,
    reportGenerated: state.reportGenerated,
    immMode: state.immDisplayMode,
    errors: errorLog.slice(-10),
  };
}

window.onerror = function(message, source, lineno, colno, error) {
  captureError({
    type: 'uncaught',
    message: String(message),
    source: source || '',
    line: lineno,
    col: colno,
    stack: error?.stack || '',
  });
  return false;
};

window.addEventListener('unhandledrejection', function(e) {
  captureError({
    type: 'unhandled-promise',
    message: String(e.reason),
    stack: e.reason?.stack || '',
  });
});


/* ── 2. STATE ──────────────────────────────────────────────── */

const state = {
  rawData:null, personnel:[], colIndex:{}, parseStats:{},
  filteredResults:[], currentStats:{}, settings:{},
  reportGenerated:false, emblemBase64:null,
  immDisplayMode:'grouped', _lastDentalFlag:false,
  errorReportPending:false,
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
  tstSub:              document.getElementById('tstSub'),
  itemTstCombine:      document.getElementById('item-tstCombine'),
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
  emblemError:         document.getElementById('emblemError'),
  infoColCount:        document.getElementById('infoColCount'),
  infoColContainer:    document.getElementById('infoColContainer'),
  projectionDate:      document.getElementById('projectionDate'),
  generateBtn:         document.getElementById('generateBtn'),
  generateBlockMsg:    document.getElementById('generateBlockMsg'),
  columnCount:         document.getElementById('columnCount'),
  columnLabel:         document.getElementById('columnLabel'),
  columnCounter:       document.getElementById('columnCounter'),
  exportSettingsBtn:   document.getElementById('exportSettingsBtn'),
  importSettingsInput: document.getElementById('importSettingsInput'),
  previewPlaceholder:  document.getElementById('previewPlaceholder'),
  reportOutput:        document.getElementById('reportOutput'),
  printOutput:         document.getElementById('printOutput'),
  exportBar:           document.getElementById('exportBar'),
  exportPdfBtn:        document.getElementById('exportPdfBtn'),
  exportHint:          document.getElementById('exportHint'),
  exportPdfOfficerBtn: document.getElementById('exportPdfOfficerBtn'),
  exportPdfEnlistedBtn:document.getElementById('exportPdfEnlistedBtn'),
  exportBtnDivider:    document.getElementById('exportBtnDivider'),
  floatingTip:         document.getElementById('floatingTip'),
  // Header pills
  feedbackBtn:         document.getElementById('feedbackBtn'),
  tutorialBtn:         document.getElementById('tutorialBtn'),
  // Feedback modal
  feedbackOverlay:     document.getElementById('feedbackOverlay'),
  feedbackClose:       document.getElementById('feedbackClose'),
  feedbackCategories:  document.getElementById('feedbackCategories'),
  feedbackText:        document.getElementById('feedbackText'),
  feedbackEmail:       document.getElementById('feedbackEmail'),
  feedbackCharCount:   document.getElementById('feedbackCharCount'),
  feedbackPrivacy:     document.getElementById('feedbackPrivacy'),
  feedbackEmailHint:   document.getElementById('feedbackEmailHint'),
  feedbackSubmit:      document.getElementById('feedbackSubmit'),
  // Error bar
  errorBar:            document.getElementById('errorBar'),
  // Feedback diagnostics
  feedbackDiagBanner:  document.getElementById('feedbackDiagBanner'),
  feedbackDiagDetails: document.getElementById('feedbackDiagDetails'),
  feedbackDiagPre:     document.getElementById('feedbackDiagPre'),
  // About modal
  versionBadge:        document.getElementById('versionBadge'),
  aboutModal:          document.getElementById('aboutModal'),
  aboutClose:          document.getElementById('aboutClose'),
  aboutContactLink:    document.getElementById('aboutContactLink'),
  aboutTabCopy:        document.getElementById('aboutTabCopy'),
  aboutTabVer:         document.getElementById('aboutTabVer'),
  aboutPanelCopy:      document.getElementById('aboutPanelCopy'),
  aboutPanelVer:       document.getElementById('aboutPanelVer'),
  // Walkthrough
  wtSpotlight:         document.getElementById('wtSpotlight'),
  wtWelcomeOverlay:    document.getElementById('wtWelcomeOverlay'),
  wtBtnStart:          document.getElementById('wtBtnStart'),
  wtCallout:           document.getElementById('wtCallout'),
  wtStepBadge:         document.getElementById('wtStepBadge'),
  wtStepOf:            document.getElementById('wtStepOf'),
  wtProgressFill:      document.getElementById('wtProgressFill'),
  wtTitle:             document.getElementById('wtTitle'),
  wtDesc:              document.getElementById('wtDesc'),
  wtTip:               document.getElementById('wtTip'),
  wtBtnSkip:           document.getElementById('wtBtnSkip'),
  wtBtnPrev:           document.getElementById('wtBtnPrev'),
  wtBtnNext:           document.getElementById('wtBtnNext'),
  wtDoneOverlay:       document.getElementById('wtDoneOverlay'),
  wtDontShow:          document.getElementById('wtDontShow'),
  wtBtnDone:           document.getElementById('wtBtnDone'),
};


/* ── 4. INITIALIZATION ─────────────────────────────────────── */

document.addEventListener('DOMContentLoaded', () => {
  console.log(`Hitlist Hatcher ${APP_VERSION} — initializing`);
  initDisclaimer();
  initSettingsPanel();
  loadSettings();
  wireUploadHandlers();
  wireExportPdfHandler();
  wireSettingsHandlers();
  wireTooltips();
  setDefaultProjectionDate();
  refreshColumnCounter();
  initFeedbackModal();
  initAboutModal();
  console.log('Ready.');
});


/* ── 4a. DISCLAIMER ────────────────────────────────────────── */

function initDisclaimer() {
  if (localStorage.getItem(DISCLAIMER_KEY) === 'true') {
    dom.disclaimerModal.classList.add('hidden');
    initWalkthrough();
    return;
  }
  dom.disclaimerModal.classList.remove('hidden');
  dom.disclaimerAccept.addEventListener('click', () => {
    if (dom.disclaimerRemember.checked) {
      localStorage.setItem(DISCLAIMER_KEY, 'true');
    }
    dom.disclaimerModal.classList.add('hidden');
    initWalkthrough();
  });
}


/* ── 4b. TOOLTIPS ──────────────────────────────────────────── */

function wireTooltips() {
  const tip = dom.floatingTip;
  const PAD = 16;
  let active = null;

  document.querySelectorAll('.tip').forEach(trigger => {
    trigger.addEventListener('mouseenter', e => {
      active = trigger;
      const title = trigger.dataset.tipTitle || '';
      const body  = trigger.dataset.tipBody  || '';
      tip.innerHTML = `<strong>${title}</strong>${body}`;
      tip.classList.add('visible');
      place(e);
    });
    trigger.addEventListener('mousemove', e => {
      if (active === trigger) place(e);
    });
    trigger.addEventListener('mouseleave', () => {
      active = null;
      tip.classList.remove('visible');
    });
  });

  function place(e) {
    let x = e.clientX + PAD;
    let y = e.clientY + PAD;
    tip.style.left = x + 'px';
    tip.style.top  = y + 'px';
    const rect = tip.getBoundingClientRect();
    if (rect.right  > window.innerWidth  - 8) x = e.clientX - rect.width  - PAD;
    if (rect.bottom > window.innerHeight - 8) y = e.clientY - rect.height - PAD;
    tip.style.left = x + 'px';
    tip.style.top  = y + 'px';
  }
}


/* ── 4c. ERROR NOTIFICATION BAR ───────────────────────────── */

const ERROR_BAR_TITLES = {
  'generate': 'Report generation failed',
  'export':   'PDF export failed',
};

function showErrorBar(errorType) {
  if(!dom.errorBar) return;
  const title = ERROR_BAR_TITLES[errorType] || 'An error occurred';
  document.getElementById('errorBarTitle').textContent = title;
  dom.errorBar.classList.remove('hidden');

  // Wire buttons (clone to avoid duplicate listeners)
  const reportBtn  = document.getElementById('errorBarReport');
  const dismissBtn = document.getElementById('errorBarDismiss');
  const closeBtn   = document.getElementById('errorBarClose');
  const newReport  = reportBtn.cloneNode(true);
  const newDismiss = dismissBtn.cloneNode(true);
  const newClose   = closeBtn.cloneNode(true);
  reportBtn.replaceWith(newReport);
  dismissBtn.replaceWith(newDismiss);
  closeBtn.replaceWith(newClose);

  newReport.addEventListener('click', reportBugFromError);
  newDismiss.addEventListener('click', dismissErrorBar);
  newClose.addEventListener('click', dismissErrorBar);
}

function dismissErrorBar() {
  if(dom.errorBar) dom.errorBar.classList.add('hidden');
}

function reportBugFromError() {
  dismissErrorBar();
  state.errorReportPending = true;
  openFeedbackWithDiagnostics();
}

function openFeedbackWithDiagnostics() {
  if(!dom.feedbackOverlay) return;
  // Pre-select Bug Report category
  dom.feedbackCategories.querySelectorAll('.cat-pill').forEach(p =>
    p.classList.toggle('selected', p.dataset.cat === 'bug'));
  // Show diagnostics banner
  if(dom.feedbackDiagBanner) dom.feedbackDiagBanner.classList.remove('hidden');
  // Populate and show diagnostics details
  if(dom.feedbackDiagDetails && dom.feedbackDiagPre) {
    dom.feedbackDiagPre.textContent = JSON.stringify(getErrorReport(), null, 2);
    dom.feedbackDiagDetails.classList.remove('hidden');
  }
  // Update placeholder and footer
  dom.feedbackText.placeholder = 'What were you doing when the error occurred? (e.g., "I clicked Generate Hit List with 12 immunizations selected in Individual mode")';
  dom.feedbackPrivacy.textContent = '🔧 Error diagnostics attached';
  // Open
  dom.feedbackOverlay.classList.remove('hidden');
}


/* ── 4d. FEEDBACK MODAL ───────────────────────────────────── */

function initFeedbackModal() {
  if(!dom.feedbackBtn) return;

  // Open / close
  dom.feedbackBtn.addEventListener('click', () => {
    clearDiagnosticsUI();
    dom.feedbackOverlay.classList.remove('hidden');
  });
  dom.feedbackClose.addEventListener('click', closeFeedback);
  dom.feedbackOverlay.addEventListener('click', e => { if(e.target===dom.feedbackOverlay) closeFeedback(); });

  // Category pills
  dom.feedbackCategories.querySelectorAll('.cat-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      dom.feedbackCategories.querySelectorAll('.cat-pill').forEach(p => p.classList.remove('selected'));
      pill.classList.add('selected');
    });
  });

  // Char counter
  dom.feedbackText.addEventListener('input', () => {
    dom.feedbackCharCount.textContent = dom.feedbackText.value.length;
  });

  // Email field — live privacy note
  dom.feedbackEmail.addEventListener('input', () => {
    const val = dom.feedbackEmail.value.trim();
    if(val) {
      dom.feedbackEmail.classList.add('has-value');
      dom.feedbackEmailHint.textContent = 'A reply will be sent to this address if needed.';
      dom.feedbackEmailHint.classList.add('has-value');
      dom.feedbackPrivacy.textContent = state.errorReportPending
        ? '🔧 Diagnostics attached · 📧 Reply to: ' + val
        : '📧 Reply to: ' + val;
    } else {
      dom.feedbackEmail.classList.remove('has-value');
      dom.feedbackEmailHint.textContent = 'Leave blank to submit anonymously.';
      dom.feedbackEmailHint.classList.remove('has-value');
      dom.feedbackPrivacy.textContent = state.errorReportPending
        ? '🔧 Error diagnostics attached'
        : '🔒 Submitted anonymously';
    }
  });

  // Submit
  dom.feedbackSubmit.addEventListener('click', async () => {
    const text = dom.feedbackText.value.trim();
    if(!text) { dom.feedbackText.focus(); return; }
    const category = dom.feedbackCategories.querySelector('.cat-pill.selected')?.dataset.cat || 'other';
    const email    = dom.feedbackEmail.value.trim();
    dom.feedbackSubmit.textContent = 'Sending…';
    dom.feedbackSubmit.disabled    = true;
    // Attach diagnostics if error report or any errors have been captured
    const payload = {category, message:text, _replyto:email||undefined};
    if(state.errorReportPending || errorLog.length > 0) {
      payload._diagnostics = JSON.stringify(getErrorReport());
    }
    try {
      await fetch('https://formspree.io/f/mvzwopqb', {
        method:'POST',
        headers:{'Content-Type':'application/json', 'Accept':'application/json'},
        body: JSON.stringify(payload),
      });
      dom.feedbackSubmit.textContent = '✓ Sent!';
      setTimeout(() => { closeFeedback(); resetFeedback(); }, 1200);
    } catch(err) {
      dom.feedbackSubmit.textContent = 'Send Feedback';
      dom.feedbackSubmit.disabled    = false;
    }
  });
}

function closeFeedback() {
  dom.feedbackOverlay.classList.add('hidden');
}

// Clears diagnostics-specific UI without clearing user text/email fields.
// Called when opening feedback normally (💬 button) and after submit (via resetFeedback).
function clearDiagnosticsUI() {
  state.errorReportPending = false;
  dom.feedbackText.placeholder = 'Describe the bug, feature, or idea…';
  dom.feedbackPrivacy.textContent = '🔒 Submitted anonymously';
  if(dom.feedbackDiagBanner)  dom.feedbackDiagBanner.classList.add('hidden');
  if(dom.feedbackDiagDetails) dom.feedbackDiagDetails.classList.add('hidden');
}

function resetFeedback() {
  dom.feedbackText.value  = '';
  dom.feedbackEmail.value = '';
  dom.feedbackCharCount.textContent = '0';
  dom.feedbackEmail.classList.remove('has-value');
  dom.feedbackEmailHint.textContent = 'Leave blank to submit anonymously.';
  dom.feedbackEmailHint.classList.remove('has-value');
  dom.feedbackSubmit.textContent  = 'Send Feedback';
  dom.feedbackSubmit.disabled     = false;
  dom.feedbackCategories.querySelectorAll('.cat-pill').forEach((p,i) => p.classList.toggle('selected', i===0));
  clearDiagnosticsUI();
}




/* ── 4e. ABOUT MODAL ───────────────────────────────────────── */

function initAboutModal() {
  // Set copyright year dynamically in both export bar and about modal
  const year = new Date().getFullYear();
  document.querySelectorAll('#copyrightYear, .about-year').forEach(el => {
    el.textContent = year;
  });

  if(!dom.versionBadge || !dom.aboutModal) return;

  function closeAbout() { dom.aboutModal.classList.add('hidden'); }

  function switchAboutTab(tab, panel) {
    [dom.aboutTabCopy, dom.aboutTabVer].forEach(t => t.classList.remove('active'));
    [dom.aboutPanelCopy, dom.aboutPanelVer].forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    panel.classList.add('active');
  }

  dom.aboutTabCopy.addEventListener('click', () => switchAboutTab(dom.aboutTabCopy, dom.aboutPanelCopy));
  dom.aboutTabVer.addEventListener('click',  () => switchAboutTab(dom.aboutTabVer,  dom.aboutPanelVer));

  // Reset to Copyright tab each time modal opens
  function openAbout() {
    switchAboutTab(dom.aboutTabCopy, dom.aboutPanelCopy);
    dom.aboutModal.classList.remove('hidden');
  }

  dom.versionBadge.addEventListener('click', openAbout);
  dom.versionBadge.addEventListener('keydown', e => {
    if(e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openAbout(); }
  });
  dom.aboutClose.addEventListener('click', closeAbout);
  dom.aboutModal.addEventListener('click', e => {
    if(e.target === dom.aboutModal) closeAbout();
  });
  if(dom.aboutContactLink) {
    dom.aboutContactLink.addEventListener('click', e => {
      e.preventDefault();
      closeAbout();
      dom.feedbackOverlay.classList.remove('hidden');
    });
  }
}

/* ── 4d. WALKTHROUGH ───────────────────────────────────────── */

const WT_STEPS = [
  {
    title: 'Upload Your MRRS File',
    desc:  'Start by uploading your <strong>Excel IMR Report (Details)</strong> export from MRRS. Drag and drop your file here, or click to browse. The file never leaves your browser.',
    tip:   '💡 Only .xlsx and .xls files are accepted. In MRRS: Reports → IMR → Force → IMR, under Report Title select Excel IMR Report (Details), within Unit select your command, then click Run.',
    target:'uploadCard',   side:'right', scrollTo:'uploadCard',
  },
  {
    title: 'Basic Report Items',
    desc:  'Select which medical readiness items to include. <strong>Basic items</strong> — PHA, Dental, HIV Lab, and Audiogram — are checked by default. Toggle any item on or off.',
    tip:   '💡 The Dental item has a sub-option to choose between a 12-month (365-day) due date or the MRRS-supplied date.',
    target:'basicItemsGroup', side:'right', scrollTo:'reportItemsCard',
  },
  {
    title: 'Immunization Options',
    desc:  '<strong>Grouped Column</strong> shows a single "Immunizations" cell with a count — hover to see which vaccines are due and their due dates. <strong>Individual Columns</strong> creates one column per vaccine. Watch the column counter when using individual columns.',
    tip:   '💡 Grouped mode is recommended for large units — it keeps the table compact and printable on one page.',
    target:'immSection', side:'right', scrollTo:null,
  },
  {
    title: 'Deployment & Accessions Items',
    desc:  '<strong>Deployment items</strong> include PDHA, PDHRA, and MHA assessments. <strong>Accessions items</strong> cover Blood Type, DNA, G6PD, and other entry requirements.',
    target:'deployAccessionGroup', side:'right', scrollTo:null,
  },
  {
    title: 'Display Options & Sorting',
    desc:  'Filter by <strong>Officers</strong>, <strong>Enlisted</strong>, or both — or generate separate reports for each. Sort by last name or section. Toggle the Rank/Rate and Section columns on or off to save horizontal space.',
    target:'displayCard', side:'right', scrollTo:'displayCard',
  },
  {
    title: 'Warning Windows',
    desc:  'The <strong>Yellow</strong> window flags personnel due within that many days (default: 7). The <strong>Green</strong> window flags those coming up soon (default: 30). Yellow must always be smaller than Green.',
    tip:   '💡 If Yellow ≥ Green, the Generate button locks until the values are corrected.',
    target:'displayCard', side:'right', scrollTo:'displayCard',
  },
  {
    title: 'Unit Customization',
    desc:  'Set your <strong>unit name</strong>, upload an <strong>emblem</strong>, and fill in the <strong>Header Information Text</strong> columns — clinic phone numbers, hours, or other unit specific information. This text prints on every report.',
    tip:   '💡 Text updates in the header in real time as you type — no need to regenerate.',
    target:'unitCard', side:'right', scrollTo:'unitCard',
  },
  {
    title: 'Projection Date',
    desc:  'Personnel are included on the report if their due date falls <strong>on or before this date</strong>. Extending it casts a wider net — useful for upcoming deployments. Color coding always reflects urgency from <em>today\'s date</em>.',
    target:'projectionCard', side:'right', scrollTo:'projectionCard',
  },
  {
    title: 'Column Counter & Warnings',
    desc:  'This counter estimates how many columns your report will have. It turns <strong>yellow</strong> as you approach the printable limit and <strong>red</strong> when you\'ve likely exceeded it. Reduce selected items or switch immunizations to Grouped mode to bring it down.',
    target:'columnCounter', side:'right', scrollTo:null,
  },
  {
    title: 'Save & Restore Settings',
    desc:  '<strong>Export Settings</strong> saves all your current selections — items, thresholds, unit name, and info text — so they can be restored later. <strong>Import Settings</strong> loads a previously saved configuration. Use this to share a standard configuration across your team.',
    tip:   '❗ Browser data may be cleared without notice, which will erase your saved settings. Export your settings to protect your configuration and make it portable — restore them on any computer or share them across your team for consistent reports.',
    tipStyle: 'amber',
    target:'settingsActions', side:'right', scrollTo:null,
  },
  {
    title: 'Generate the Hit List',
    desc:  'Once a file is uploaded and your settings are configured, click <strong>Generate Hit List</strong>. The report appears instantly in the preview panel — only personnel with items due within your projection window will appear.',
    target:'generateBtn', side:'right', scrollTo:null,
  },
  {
    title: 'The Generated Report',
    desc:  'The report header shows your <strong>unit name</strong>, report and projection dates, <strong>readiness percentages</strong> from MRRS, and your header information text. Each row is a person with at least one item due.',
    target:'previewPanel', side:'left', scrollTo:null,
  },
  {
    title: 'Color Coding & Legend',
    desc:  'Cells are color-coded by urgency from <strong>today\'s date</strong>: <span style="background:#e8b4ae;padding:1px 6px;border-radius:3px;color:#6a2018;font-weight:700;">Red</span> = overdue, <span style="background:#e0cc9a;padding:1px 6px;border-radius:3px;color:#584010;font-weight:700;">Yellow</span> = due soon, <span style="background:#a4c8b4;padding:1px 6px;border-radius:3px;color:#1a3e28;font-weight:700;">Green</span> = upcoming, <span style="background:#fff;padding:1px 6px;border-radius:3px;border:1px solid #c8ceda;">White</span> = within projection window but not yet urgent. Dashes (—) mean the item is not due.',
    target:'previewPanel', side:'left', scrollTo:null,
  },
  {
    title: 'Grouped Immunization Tooltip',
    desc:  'When immunizations are in <strong>Grouped mode</strong>, cells show a count like "2 due." Hover over any cell to see a tooltip listing each due vaccine alongside its due date.',
    tip:   '💡 Hover over a cell in the Immunizations column on the live report to try it.',
    target:'previewPanel', side:'left', scrollTo:null,
  },
  {
    title: 'Export to PDF',
    desc:  'Click the <strong>Export PDF</strong> button below the report. When using <strong>Separate Reports</strong>, individual buttons appear for each group. In the print dialog, set the destination to <strong>"Save as PDF."</strong> The full header repeats on every page so each printed page is self-contained.',
    target:'exportBar', side:'left', scrollTo:null,
  },
  {
    title: 'Feedback & Feature Requests',
    desc:  'The <strong>💬 Feedback</strong> button in the top-right corner is always available. Use it to report bugs, request new features, or suggest ideas for other tools that could make your workflow more efficient.',
    tip:   '💡 All feedback goes directly to the development team.',
    target:'feedbackBtn', side:'bottom', scrollTo:null,
  },
  {
    title: 'Replay This Tutorial',
    desc:  'The <strong>❓ Tutorial</strong> button is always available in the top-right corner. Click it at any time to restart this walkthrough from the beginning.',
    target:'tutorialBtn', side:'bottom', scrollTo:null,
  },
];

let wtCurrentStep = 0;

function initWalkthrough() {
  if(!dom.wtWelcomeOverlay) return;

  // Wire Tutorial button
  if(dom.tutorialBtn) dom.tutorialBtn.addEventListener('click', launchWalkthrough);

  // If already seen, skip
  if(localStorage.getItem(WALKTHROUGH_KEY) === 'true') return;

  // Show welcome screen after a short settle
  setTimeout(() => dom.wtWelcomeOverlay.classList.remove('hidden'), 120);
  wireWalkthroughHandlers();
}

function launchWalkthrough() {
  if(!dom.wtWelcomeOverlay) return;
  wtCurrentStep = 0;
  dom.wtDoneOverlay.classList.add('hidden');
  dom.wtCallout.classList.add('hidden');
  dom.wtSpotlight.classList.add('hidden');
  wireWalkthroughHandlers();
  dom.wtWelcomeOverlay.classList.remove('hidden');
}

function wireWalkthroughHandlers() {
  // Avoid duplicate listeners
  const btnStart = dom.wtBtnStart;
  const newStart = btnStart.cloneNode(true);
  btnStart.parentNode.replaceChild(newStart, btnStart);
  dom.wtBtnStart = document.getElementById('wtBtnStart');

  dom.wtBtnStart.addEventListener('click', () => {
    dom.wtWelcomeOverlay.classList.add('hidden');
    wtShowStep(0);
  });
  dom.wtBtnNext.onclick  = () => {
    if(wtCurrentStep < WT_STEPS.length - 1) wtShowStep(wtCurrentStep + 1);
    else wtEnd();
  };
  dom.wtBtnPrev.onclick  = () => { if(wtCurrentStep > 0) wtShowStep(wtCurrentStep - 1); };
  dom.wtBtnSkip.onclick  = wtEnd;
  dom.wtBtnDone.onclick  = () => {
    if(dom.wtDontShow.checked) localStorage.setItem(WALKTHROUGH_KEY, 'true');
    dom.wtDoneOverlay.classList.add('hidden');
  };
}

function wtShowStep(idx) {
  wtCurrentStep = idx;
  const step  = WT_STEPS[idx];
  const total = WT_STEPS.length;

  // Fix 4B: inject sample report when entering steps 12-14 (idx 11-13),
  // clear it when navigating away from that range.
  if(WT_SAMPLE_STEPS.includes(idx)) {
    wtInjectSampleReport();
  } else {
    wtClearSampleReport();
  }

  // Scroll the settings panel to bring the target element into view.
  // getBoundingClientRect delta is used instead of offsetTop — offsetTop
  // resolves relative to the nearest positioned ancestor which may not be
  // .settings-scroll, giving incorrect results in nested scroll containers.
  const scrollTarget = step.scrollTo
    ? document.getElementById(step.scrollTo)
    : document.getElementById(step.target);
  const settingsScroll = document.querySelector('.settings-scroll');
  if(scrollTarget && settingsScroll && settingsScroll.contains(scrollTarget)) {
    const SCROLL_PAD    = 20; // px breathing room above the spotlight ring
    const containerRect = settingsScroll.getBoundingClientRect();
    const targetRect    = scrollTarget.getBoundingClientRect();
    settingsScroll.scrollTop += (targetRect.top - containerRect.top) - SCROLL_PAD;
  }

  // Update callout content
  dom.wtTitle.innerHTML      = step.title;
  dom.wtDesc.innerHTML       = step.desc;
  dom.wtStepBadge.textContent = `Step ${idx + 1}`;
  dom.wtStepOf.textContent    = `${idx + 1} of ${total}`;
  dom.wtProgressFill.style.width = `${((idx + 1) / total) * 100}%`;

  if(step.tip) {
    dom.wtTip.innerHTML = step.tip;
    dom.wtTip.classList.remove('hidden');
    dom.wtTip.classList.toggle('wt-callout-tip-amber', step.tipStyle === 'amber');
  } else if(idx === 12 && dom.reportOutput.dataset.wtSample === 'true') {
    // Step 13 (Color Coding) has no static tip — show sample note instead
    dom.wtTip.innerHTML = '💡 Sample report shown for illustration — generate with your MRRS file to see real data.';
    dom.wtTip.classList.remove('hidden');
    dom.wtTip.classList.remove('wt-callout-tip-amber');
  } else {
    dom.wtTip.classList.add('hidden');
    dom.wtTip.classList.remove('wt-callout-tip-amber');
  }

  dom.wtBtnPrev.disabled    = (idx === 0);
  dom.wtBtnNext.textContent = (idx === total - 1) ? 'Finish ✓' : 'Next →';

  dom.wtCallout.classList.remove('hidden');
  dom.wtSpotlight.classList.remove('hidden');

  requestAnimationFrame(() => {
    wtPositionSpotlight(step);
    wtPositionCallout(step);
  });
}

function wtPositionSpotlight(step) {
  const PAD = 8;
  const el  = document.getElementById(step.target);
  if(!el) return;
  const r = el.getBoundingClientRect();

  // Raw spotlight bounds
  let top    = r.top    - PAD;
  let left   = r.left   - PAD;
  let right  = r.right  + PAD;
  let bottom = r.bottom + PAD;

  // Clamp to viewport, inset by SHADOW_W so the 4px box-shadow ring
  // doesn't paint beyond the visible content area edge.
  const SHADOW_W = 4;
  top    = Math.max(top,    SHADOW_W);
  left   = Math.max(left,   SHADOW_W);
  right  = Math.min(right,  document.documentElement.clientWidth  - SHADOW_W);
  bottom = Math.min(bottom, document.documentElement.clientHeight - SHADOW_W);

  // For targets inside the settings scroll area, clamp bottom against the
  // scroll container so the spotlight doesn't bleed over the sticky footer.
  const settingsScroll = document.querySelector('.settings-scroll');
  if(settingsScroll && settingsScroll.contains(el)) {
    const scrollRect = settingsScroll.getBoundingClientRect();
    bottom = Math.min(bottom, scrollRect.bottom);
  }

  dom.wtSpotlight.style.top    = top    + 'px';
  dom.wtSpotlight.style.left   = left   + 'px';
  dom.wtSpotlight.style.width  = Math.max(right  - left, 0) + 'px';
  dom.wtSpotlight.style.height = Math.max(bottom - top,  0) + 'px';
}

function wtPositionCallout(step) {
  const CALLOUT_W   = 340;
  const CALLOUT_GAP = 18;
  const el  = document.getElementById(step.target);
  if(!el) return;
  const r = el.getBoundingClientRect();
  dom.wtCallout.style.transform = '';

  if(step.side === 'right') {
    dom.wtCallout.style.left = (r.right + CALLOUT_GAP) + 'px';
    dom.wtCallout.style.top  = Math.min(Math.max(r.top, 60), window.innerHeight - 380) + 'px';
  } else if(step.side === 'left') {
    dom.wtCallout.style.left = Math.max(r.left - CALLOUT_W - CALLOUT_GAP, 10) + 'px';
    dom.wtCallout.style.top  = Math.min(Math.max(r.top, 60), window.innerHeight - 380) + 'px';
  } else if(step.side === 'bottom') {
    const x = r.left + r.width/2 - CALLOUT_W/2;
    dom.wtCallout.style.left = Math.max(Math.min(x, window.innerWidth - CALLOUT_W - 10), 10) + 'px';
    dom.wtCallout.style.top  = (r.bottom + CALLOUT_GAP) + 'px';
  }
}

function wtEnd() {
  dom.wtCallout.classList.add('hidden');
  dom.wtSpotlight.classList.add('hidden');
  wtClearSampleReport();
  dom.wtDoneOverlay.classList.remove('hidden');
}

// ── Walkthrough sample report (Fix 4B) ─────────────────────
// Injected when reaching step 12 with no real report generated.
// Cleaned up when leaving steps 11-13 backwards or on tour end.

const WT_SAMPLE_STEPS = [11, 12, 13]; // 0-based indices of steps 12-14

function wtInjectSampleReport() {
  if(state.reportGenerated) return;  // real report already present
  if(dom.reportOutput.dataset.wtSample === 'true') return; // already injected

  const today = new Date();
  // Helper: build a date relative to today
  function relDate(days) {
    const d = new Date(today); d.setDate(d.getDate() + days); return d;
  }
  function fmt(d) {
    const m = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    return `${d.getDate()} ${m[d.getMonth()]}`;
  }
  function fmtFull(d) {
    const m = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    return `${d.getDate()} ${m[d.getMonth()]} ${d.getFullYear()}`;
  }

  // Build a tiny tooltip payload for the immunization cell
  const immItems = [
    {label:'Influenza', due: relDate(12).getTime()},
    {label:'TDap',      due: relDate(45).getTime()},
  ];
  const immJson = escHtml(JSON.stringify(immItems));

  const reportDate = fmtFull(today);
  const projDate   = fmtFull(relDate(55));

  const sampleHTML = `
<div class="hitlist-wrapper">
  <div class="hitlist-header" style="position:relative;">
    <div class="hitlist-header-top">
      <div class="hitlist-emblem-placeholder">Unit<br/>Emblem</div>
      <div class="hitlist-title-block">
        <div class="hitlist-unit-name">Sample Unit Medical Hit List</div>
        <div class="hitlist-title-line">Report Date: ${reportDate}</div>
        <div class="hitlist-projection-line">Projected to: ${projDate}</div>
      </div>
      <div class="hitlist-emblem-placeholder">Unit<br/>Emblem</div>
    </div>
    <div class="hitlist-stats-bar">
      <div class="hitlist-stat"><span class="hitlist-stat-label">Fully Ready</span><span class="hitlist-stat-value stat-good">88.0%</span></div>
      <div class="hitlist-stat"><span class="hitlist-stat-label">Not Ready</span><span class="hitlist-stat-value stat-bad">5.0%</span></div>
      <div class="hitlist-stat"><span class="hitlist-stat-label">Partial</span><span class="hitlist-stat-value stat-warn">4.0%</span></div>
      <div class="hitlist-stat"><span class="hitlist-stat-label">Indeterminate</span><span class="hitlist-stat-value stat-neutral">3.0%</span></div>
    </div>
  </div>
  <div class="hitlist-legend">
    <span class="hitlist-legend-label">Key:</span>
    <span class="legend-item"><span class="legend-swatch red"></span>Overdue / Class 3–4</span>
    <span class="legend-item"><span class="legend-swatch yellow"></span>Due within 7 days</span>
    <span class="legend-item"><span class="legend-swatch green"></span>Due within 30 days</span>
    <span class="legend-item"><span class="legend-swatch none" style="border:1px solid #c8ceda;"></span>Due beyond 30 days</span>
  </div>
  <div class="hitlist-table-wrapper" style="margin-top:4px;">
    <table class="hitlist-table">
      <thead><tr>
        <th class="col-name-hdr">Name</th>
        <th>Rank</th>
        <th>PHA</th>
        <th>HIV Lab</th>
        <th>Audiogram</th>
        <th>Immunizations</th>
      </tr></thead>
      <tbody>
        <tr>
          <td class="col-name">Anderson James</td>
          <td class="col-rank">SGT</td>
          <td class="col-item status-red">${fmt(relDate(-12))}</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
        </tr>
        <tr>
          <td class="col-name">Chen Mei</td>
          <td class="col-rank">CPL</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item status-yellow">${fmt(relDate(4))}</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
        </tr>
        <tr>
          <td class="col-name">Davis Robert</td>
          <td class="col-rank">SSGT</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item status-green">${fmt(relDate(18))}</td>
          <td class="col-item cell-na">—</td>
        </tr>
        <tr>
          <td class="col-name">Garcia Elena</td>
          <td class="col-rank">LCPL</td>
          <td class="col-item">${fmt(relDate(40))}</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
        </tr>
        <tr>
          <td class="col-name">Kim Jason</td>
          <td class="col-rank">CPL</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item status-green"><div class="imm-grouped" data-tooltip="${immJson}">2 due</div></td>
        </tr>
        <tr>
          <td class="col-name">Miller Sarah</td>
          <td class="col-rank">SGT</td>
          <td class="col-item status-red">${fmt(relDate(-3))}</td>
          <td class="col-item status-red">${fmt(relDate(-30))}</td>
          <td class="col-item cell-na">—</td>
          <td class="col-item cell-na">—</td>
        </tr>
      </tbody>
    </table>
  </div>
</div>`;

  dom.previewPlaceholder.classList.add('hidden');
  dom.reportOutput.innerHTML = sampleHTML;
  dom.reportOutput.dataset.wtSample = 'true';
  wireImmTooltips();
}

function wtClearSampleReport() {
  if(dom.reportOutput.dataset.wtSample !== 'true') return;
  dom.reportOutput.innerHTML = '';
  delete dom.reportOutput.dataset.wtSample;
  if(!state.reportGenerated) dom.previewPlaceholder.classList.remove('hidden');
}

window.addEventListener('resize', () => {
  if(!dom.wtCallout || dom.wtCallout.classList.contains('hidden')) return;
  wtPositionSpotlight(WT_STEPS[wtCurrentStep]);
  wtPositionCallout(WT_STEPS[wtCurrentStep]);
});


/* ── 5. FILE UPLOAD ────────────────────────────────────────── */

function wireUploadHandlers() {
  dom.uploadArea.addEventListener('click', e => {
    // Prevent double file dialog: the <label for="fileInput"> natively opens the
    // picker, so skip the programmatic .click() when the label or input is clicked.
    if(e.target === dom.fileInput || e.target.closest('label[for="fileInput"]')) return;
    dom.fileInput.click();
  });
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
    let wb;
    try {
      wb = XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true,raw:false});
      const formatCheck = validateMrrsFormat(wb);
      if(!formatCheck.valid){
        setUploadStatus('error', MRRS_FORMAT_ERROR);
        dom.uploadArea.classList.add('upload-error'); dom.uploadArea.classList.remove('upload-success');
        console.warn('MRRS format validation failed:', formatCheck.reason);
        return;
      }
      state.rawData = wb;
      const parsed  = parseMRRS(wb, false, formatCheck.rows);
      if (!parsed.personnel.length) {
        setUploadStatus('error','✗ No personnel rows found.'); dom.uploadArea.classList.add('upload-error'); return;
      }
      state.personnel = parsed.personnel; state.colIndex = parsed.colIndex; state.parseStats = parsed.stats;
      dom.uploadArea.classList.add('upload-success'); dom.uploadArea.classList.remove('upload-error');
      setUploadStatus('success',
        `✓ ${file.name} — ${parsed.personnel.length} personnel loaded `+
        `(${parsed.stats.officers} officers, ${parsed.stats.enlisted} enlisted)`);
      dom.generateBtn.disabled = false;
      updateGenerateBtnState();
      logParseResults(parsed);
      refreshColumnCounter();
    } catch(err) {
      captureError({
        type: 'parse-error',
        message: err.message,
        stack: err.stack || '',
        context: {
          fileName: file.name,
          fileSize: file.size,
          sheetsFound: wb ? wb.SheetNames : [],
        },
      });
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

// Validates that the uploaded workbook matches the expected MRRS IMR Detail format.
// Returns { valid:true } or { valid:false, reason:string }.
function validateMrrsFormat(workbook) {
  const sheet = workbook.Sheets[MRRS_SHEET_NAME];
  if(!sheet) return { valid:false, reason:`Sheet "${MRRS_SHEET_NAME}" not found.` };

  const rows = XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false});

  // Step A: Check that the first two rows contain FOUO / Privacy Act warning text
  const topCells = [].concat(rows[0]||[], rows[1]||[]);
  const topText  = topCells.join(' ');
  const hasFouo  = MRRS_FOUO_KEYWORDS.some(kw => topText.toUpperCase().includes(kw.toUpperCase()));
  if(!hasFouo) return { valid:false, reason:'Expected FOUO/Privacy Act warning in rows 1–2.' };

  // Step B: Check that signature columns exist in the header row
  const headerRow = (rows[MRRS_HEADER_ROW]||[]).map(v => String(v).trim());
  const missing = MRRS_SIGNATURE_COLUMNS.filter(col => !headerRow.includes(col));
  if(missing.length) return { valid:false, reason:`Missing expected columns: ${missing.join(', ')}` };

  return { valid:true, rows };
}

function parseMRRS(workbook, useMrrsDate, preloadedRows) {
  const sheet = workbook.Sheets[MRRS_SHEET_NAME];
  const rows  = preloadedRows || XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false});
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
    const m = s.match(/\*?Due\s+(\d{1,2}-[A-Z]{3}-\d{4})/i);
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
        influenza:  {due:parseDate(getCell(row,'INFLUENZA_DUE')),   label:'Influenza'},
        tdap:       {due:parseDate(getCell(row,'TDAP_DUE')),        label:'TDap'},
        typhoid:    {due:parseDate(getCell(row,'TYPHOID_DUE')),     label:'Typhoid'},
        varicella:  {due:parseDate(getCell(row,'VARICELLA_DUE')),   label:'Varicella'},
        mmr:        {due:parseDate(getCell(row,'MMR_DUE')),         label:'MMR'},
        hepa:       {due:parseDate(getCell(row,'HEPA_DUE')),        label:'Hep A'},
        hepb:       {due:parseDate(getCell(row,'HEPB_DUE')),        label:'Hep B'},
        twinrix:    {due:parseDate(getCell(row,'TWINRIX_DUE')),     label:'TwinRix'},
        rabies:     {due:parseDate(getCell(row,'RABIES_DUE')),      label:'Rabies'},
        rabiesTiter:{due:parseDate(getCell(row,'RABIES_TITER_DUE')),label:'Rabies Titer'},
        cholera:    {due:parseDate(getCell(row,'CHOLERA_DUE')),     label:'Cholera'},
        jev:        {due:parseDate(getCell(row,'JEV_DUE')),         label:'JEV'},
        mgc:        {due:parseDate(getCell(row,'MGC_DUE')),         label:'MGC'},
        polio:      {due:parseDate(getCell(row,'POLIO_DUE')),       label:'Polio'},
        yellowFever:{due:parseDate(getCell(row,'YF_DUE')),          label:'Yellow Fever'},
        anthrax:    {due:parseDate(getCell(row,'ANTHRAX_DUE')),     label:'Anthrax'},
        smallpox:   {due:parseDate(getCell(row,'SMALLPOX_DUE')),    label:'Smallpox'},
        adenovirus: {due:parseDate(getCell(row,'ADENOVIRUS_DUE')),  label:'Adenovirus'},
        pneumo:     {due:parseDate(getCell(row,'PNEUMO_DUE')),      label:'Pneumococcal'},
        hpv:        {due:parseDate(getCell(row,'HPV_DUE')),         label:'HPV'},
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

// Determine color status based on today's date only.
// today = new Date() passed in so all cells in one report share the same instant.
function colorStatus(dueDate, today, thresholds) {
  if(!dueDate||!(dueDate instanceof Date)||isNaN(dueDate)) return STATUS.NA;
  const days = daysDiff(dueDate, today);
  if(days <= 0)                  return STATUS.OVERDUE;
  if(days <= thresholds.yellow)  return STATUS.DUE_SOON;
  if(days <= thresholds.green)   return STATUS.UPCOMING;
  return STATUS.OK;
}

// Is this due date on or before the projection date?
function isDueByProjDate(dueDate, projDate) {
  if(!dueDate||!(dueDate instanceof Date)||isNaN(dueDate)) return false;
  return dueDate <= projDate;
}

function evaluateDental(person, today, projDate, thresholds) {
  if(person.dentalClass===3) return {status:STATUS.OVERDUE, displayText:'Class 3', includeInReport:true};
  if(person.dentalClass===4) return {status:STATUS.OVERDUE, displayText:'Class 4', includeInReport:true};
  const status = colorStatus(person.dentalDue, today, thresholds);
  const include = isDueByProjDate(person.dentalDue, projDate);
  return {status, displayText:null, includeInReport:include};
}

function evaluateMHA(mha, today, projDate, thresholds) {
  if(!mha || mha.naFlag) return {status:STATUS.NA, displayText:null, includeInReport:false};
  if(mha.ok)             return {status:STATUS.OK, displayText:null, includeInReport:false};
  if(mha.due) {
    const status  = colorStatus(mha.due, today, thresholds);
    const include = isDueByProjDate(mha.due, projDate);
    return {status, displayText:null, includeInReport:include};
  }
  // *Due present but date failed to parse — always show
  return {status:STATUS.OVERDUE, displayText:'Due', includeInReport:true};
}

function evaluateBloodType(bloodType) {
  if(bloodType && bloodType.trim()!=='') return {status:STATUS.OK, displayText:bloodType.trim(), includeInReport:false};
  return {status:STATUS.OVERDUE, displayText:'Missing', includeInReport:true};
}

// Simple date items (PHA, HIV, etc): color from today, inclusion from projDate
function evaluateDateItem(dueDate, today, projDate, thresholds) {
  if(!dueDate||!(dueDate instanceof Date)||isNaN(dueDate))
    return {status:STATUS.NA, displayText:null, includeInReport:false};
  const status  = colorStatus(dueDate, today, thresholds);
  const include = isDueByProjDate(dueDate, projDate);
  return {status, displayText:null, includeInReport:include};
}

// Simplified immunization: blank Due = NA (not required / deferred / complete)
function evaluateImmunization(immObj, today, projDate, thresholds) {
  if(!immObj.due) return {status:STATUS.NA, includeInReport:false};
  const status  = colorStatus(immObj.due, today, thresholds);
  const include = isDueByProjDate(immObj.due, projDate);
  return {status, includeInReport:include};
}

function evaluateImmunizations(person, selectedKeys, today, projDate, thresholds) {
  let worstStatus = STATUS.NA;
  let count = 0, includeGrouped = false;
  const dueItems = [], perVaccine = {};
  selectedKeys.forEach(key => {
    const imm = person.immunizations[key]; if(!imm) return;
    const res = evaluateImmunization(imm, today, projDate, thresholds);
    perVaccine[key] = {status:res.status, due:imm.due, includeInReport:res.includeInReport};
    if(SEVERITY[res.status] > SEVERITY[worstStatus]) worstStatus = res.status;
    if(res.includeInReport) {
      includeGrouped = true;
      count++;
      dueItems.push({label:imm.label, due:imm.due});
    }
  });
  return {worstStatus, count, dueItems, perVaccine, includeGrouped};
}

function evaluatePerson(person, settings, today, projDate) {
  const t = settings.thresholds || DEFAULT_THRESHOLDS;
  const na = {status:STATUS.NA, displayText:null, includeInReport:false};

  const phaResult           = settings.items.pha           ? evaluateDateItem(person.phaDue,         today, projDate, t) : na;
  const dentalResult        = settings.items.dental         ? evaluateDental(person,                  today, projDate, t) : na;
  const hivResult           = settings.items.hiv            ? evaluateDateItem(person.hivDue,          today, projDate, t) : na;
  const audioResult         = settings.items.audio          ? evaluateDateItem(person.audioDue,        today, projDate, t) : na;
  const pdhaResult          = settings.items.pdha           ? evaluateDateItem(person.pdhaDue,         today, projDate, t) : na;
  const pdhraResult         = settings.items.pdhra          ? evaluateDateItem(person.pdhraDue,        today, projDate, t) : na;
  const mha2Result          = settings.items.mha2           ? evaluateMHA(person.mha2,                 today, projDate, t) : na;
  const mha3Result          = settings.items.mha3           ? evaluateMHA(person.mha3,                 today, projDate, t) : na;
  const mha4Result          = settings.items.mha4           ? evaluateMHA(person.mha4,                 today, projDate, t) : na;
  const warningTagResult    = settings.items.warningTag     ? evaluateDateItem(person.warningTagDue,   today, projDate, t) : na;
  const verifyGlassesResult = settings.items.verifyGlasses  ? evaluateDateItem(person.verifyGlassesDue,today, projDate, t) : na;
  const verifyInsertsResult = settings.items.verifyInserts   ? evaluateDateItem(person.verifyInsertsDue,today, projDate, t) : na;
  const bloodTypeResult     = settings.items.bloodType      ? evaluateBloodType(person.bloodType)     : na;
  const refAudioResult      = settings.items.refAudio       ? evaluateDateItem(person.refAudioDue,    today, projDate, t) : na;
  const dnaResult           = settings.items.dna            ? evaluateDateItem(person.dnaDue,         today, projDate, t) : na;
  const g6pdResult          = settings.items.g6pd           ? evaluateDateItem(person.g6pdDue,        today, projDate, t) : na;
  const sickleResult        = settings.items.sickle         ? evaluateDateItem(person.sickleDue,      today, projDate, t) : na;
  const tstResult      = settings.items.tst      ? evaluateDateItem(person.tstDue,      today, projDate, t) : na;
  const tstQuestResult = settings.items.tstQuest ? evaluateDateItem(person.tstQuestDue, today, projDate, t) : na;

  // TST combined: pick worst status; include if either is due within projDate
  let tstCombinedResult = na;
  if(settings.items.tstCombine) {
    const tr = evaluateDateItem(person.tstDue,      today, projDate, t);
    const qr = evaluateDateItem(person.tstQuestDue, today, projDate, t);
    const worstStatus  = SEVERITY[tr.status] >= SEVERITY[qr.status] ? tr.status : qr.status;
    const include      = tr.includeInReport || qr.includeInReport;
    // Store individual items for tooltip
    const tooltipItems = [];
    if(tr.includeInReport) tooltipItems.push({label:'TST Due',   due:person.tstDue});
    if(qr.includeInReport) tooltipItems.push({label:'TST Quest', due:person.tstQuestDue});
    tstCombinedResult = {status:worstStatus, displayText:null, includeInReport:include, tooltipItems};
  }

  let wellWomanResult = na;
  if(settings.items.wellWoman && person.sex.toLowerCase()==='female') {
    const mr = evaluateDateItem(person.mammogramDue, today, projDate, t);
    const pr = evaluateDateItem(person.papDue,       today, projDate, t);
    const status = SEVERITY[mr.status] >= SEVERITY[pr.status] ? mr.status : pr.status;
    const include = mr.includeInReport || pr.includeInReport;
    wellWomanResult = {status, displayText:null, includeInReport:include};
  }

  const immKeys = settings.immunizationKeys || [];
  const immRes  = evaluateImmunizations(person, immKeys, today, projDate, t);
  const immunizationsResult = {
    grouped:{status:immRes.worstStatus, count:immRes.count, dueItems:immRes.dueItems, includeInReport:immRes.includeGrouped},
    perVaccine:immRes.perVaccine,
  };

  // isDue: person appears on report if ANY checked item says includeInReport
  const allResults = [
    phaResult, dentalResult, hivResult, audioResult,
    pdhaResult, pdhraResult, mha2Result, mha3Result, mha4Result,
    warningTagResult, verifyGlassesResult, verifyInsertsResult,
    bloodTypeResult, refAudioResult, dnaResult, g6pdResult, sickleResult,
    tstResult, tstQuestResult, tstCombinedResult, wellWomanResult,
  ];
  const isDue = allResults.some(r => r.includeInReport) || immRes.includeGrouped;

  return {person, isDue, items:{
    pha:phaResult, dental:dentalResult, hiv:hivResult, audio:audioResult,
    pdha:pdhaResult, pdhra:pdhraResult,
    mha2:mha2Result, mha3:mha3Result, mha4:mha4Result,
    warningTag:warningTagResult, verifyGlasses:verifyGlassesResult, verifyInserts:verifyInsertsResult,
    bloodType:bloodTypeResult, refAudio:refAudioResult, dna:dnaResult, g6pd:g6pdResult, sickle:sickleResult,
    tst:tstResult, tstQuest:tstQuestResult, tstCombined:tstCombinedResult, wellWoman:wellWomanResult,
    immunizations:immunizationsResult,
  }};
}

function applyFilters(personnel, settings, today, projDate){
  const allEval = personnel.map(p => evaluatePerson(p, settings, today, projDate));
  const stats   = computeReadinessStats(personnel);
  const filter  = settings.offEnlFilter || 'combined';
  let filtered  = allEval.filter(r => {
    if(!r.isDue) return false;
    if(filter === 'officer')  return r.person.offEnl === 'Officer';
    if(filter === 'enlisted') return r.person.offEnl === 'Enlisted';
    return true;
  });
  const sortBy = settings.sortBy || 'name';
  filtered.sort((a,b) => {
    if(sortBy === 'section'){const c = a.person.section.localeCompare(b.person.section); if(c!==0) return c;}
    return a.person.name.localeCompare(b.person.name);
  });
  return {results:filtered, stats};
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

const HINT_SINGLE   = 'In the print dialog, set Destination to \u201cSave as PDF\u201d';
const HINT_SEPARATE = 'Each button exports its report independently \u2014 set Destination to \u201cSave as PDF\u201d in the print dialog';

function setExportBarMode(mode, officerEmpty, enlistedEmpty) {
  // mode: 'single' | 'separate'
  const isSep = mode === 'separate';
  dom.exportPdfBtn.classList.toggle('hidden', isSep);
  dom.exportPdfOfficerBtn.classList.toggle('hidden', !isSep);
  dom.exportBtnDivider.classList.toggle('hidden', !isSep);
  dom.exportPdfEnlistedBtn.classList.toggle('hidden', !isSep);
  if(isSep) {
    dom.exportPdfOfficerBtn.disabled  = !!officerEmpty;
    dom.exportPdfEnlistedBtn.disabled = !!enlistedEmpty;
    dom.exportHint.textContent = HINT_SEPARATE;
  } else {
    dom.exportHint.textContent = HINT_SINGLE;
  }
}

function renderReport(results, stats, settings, today, projDate){
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
    const colDefs=getColumnDefs(settings);
    const activeOff = officers.length ? getActiveColumnDefs(colDefs, officers) : colDefs;
    const activeEnl = enlisted.length ? getActiveColumnDefs(colDefs, enlisted) : colDefs;
    updateColumnCounter(Math.max(activeOff.length, activeEnl.length), true);
    let html='';
    if(officers.length) html+=buildHitListHTML(officers,stats,settings,today,projDate,activeOff,'Officer Hit List','hitlist-officer');
    if(enlisted.length) html+=`<div style="margin-top:32px;">${buildHitListHTML(enlisted,stats,settings,today,projDate,activeEnl,'Enlisted Hit List','hitlist-enlisted')}</div>`;
    dom.previewPlaceholder.classList.add('hidden');
    dom.reportOutput.innerHTML=html;
    wireImmTooltips();
    wireTstTooltip();
    setExportBarMode('separate', officers.length===0, enlisted.length===0);
    dom.exportBar.hidden=false; state.reportGenerated=true; return;
  }

  const colDefs   = getColumnDefs(settings);
  const activeColDefs = getActiveColumnDefs(colDefs, results);
  updateColumnCounter(activeColDefs.length, true);
  dom.previewPlaceholder.classList.add('hidden');
  dom.reportOutput.innerHTML=buildHitListHTML(results,stats,settings,today,projDate,activeColDefs,null,null);
  wireImmTooltips();
  wireTstTooltip();
  setExportBarMode('single');
  dom.exportBar.hidden=false; state.reportGenerated=true;
}

function buildHitListHTML(results, stats, settings, today, projDate, colDefs, subtitleOverride, wrapperId){
  const idAttr = wrapperId ? ` id="${wrapperId}"` : '';
  return`<div class="hitlist-wrapper"${idAttr}>
    ${buildHeader(stats,settings,projDate,subtitleOverride)}
    ${buildLegend(settings)}
    <div class="hitlist-table-wrapper">${buildTable(results,colDefs,settings)}</div>
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
  if(settings.items.tstCombine) {
    d.push({key:'tst-combined', label:'TST', type:'tst-combined'});
  } else {
    if(settings.items.tst)      d.push({key:'tst',      label:'TST',       type:'item'});
    if(settings.items.tstQuest) d.push({key:'tstQuest', label:'TST Quest', type:'item'});
  }
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

// Returns only columns where at least one person has a non-NA, non-OK status
// OR the column is an identity column (name/rank/section).
// Note: column counter still uses getColumnDefs (pre-suppression estimate).
function getActiveColumnDefs(colDefs, results) {
  return colDefs.filter(col => {
    if(col.type === 'identity') return true;
    if(col.type === 'imm-grouped') {
      return results.some(r => r.items.immunizations.grouped.includeInReport);
    }
    if(col.type === 'imm-individual') {
      return results.some(r => {
        const pv = r.items.immunizations.perVaccine[col.vaccKey];
        return pv && pv.includeInReport;
      });
    }
    if(col.type === 'tst-combined') {
      return results.some(r => r.items.tstCombined && r.items.tstCombined.includeInReport);
    }
    // type === 'item'
    return results.some(r => {
      const ir = r.items[col.key];
      return ir && ir.includeInReport;
    });
  });
}

function buildHeader(stats, settings, projDate, subtitleOverride){
  const unitName  =escHtml(settings.unitName||'Unit Name');
  const projStr   =formatDateFull(projDate);
  const reportDate=formatDateFull(new Date());

  const emblem=settings.emblemBase64
    ?`<img class="hitlist-emblem" src="${settings.emblemBase64}" alt="Unit Emblem" />`
    :`<div class="hitlist-emblem-placeholder">Unit<br/>Emblem</div>`;

  const titleLine = `<div class="hitlist-title-line">Report Date: ${reportDate}</div>`;
  const subtitleLine = subtitleOverride
    ? `<div class="hitlist-subtitle-line">${escHtml(subtitleOverride)}</div>`
    : '';

  // Multi-column info bar
  const infoBar = buildInfoBar(settings);

  return`<div class="hitlist-header">
    <div class="hitlist-header-top">${emblem}
      <div class="hitlist-title-block">
        <div class="hitlist-unit-name">${unitName} Medical Hit List</div>
        ${titleLine}
        <div class="hitlist-projection-line">Projected to: ${projStr}</div>
        ${subtitleLine}
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
  const gridCols=`repeat(${visibleCols.length},1fr)`;
  return`<div class="hitlist-info-bar hitlist-info-grid" style="grid-template-columns:${gridCols}">${cells}</div>`;
}

function buildLegend(settings){
  const t = settings.thresholds || DEFAULT_THRESHOLDS;
  return`<div class="hitlist-legend">
    <span class="hitlist-legend-label">Key:</span>
    <span class="legend-item"><span class="legend-swatch red"></span>Overdue / Class 3–4</span>
    <span class="legend-item"><span class="legend-swatch yellow"></span>Due within ${t.yellow} days</span>
    <span class="legend-item"><span class="legend-swatch green"></span>Due within ${t.green} days</span>
    <span class="legend-item"><span class="legend-swatch none" style="border:1px solid #c8ceda;"></span>Due beyond ${t.green} days</span>
  </div>`;
}

function buildTable(results, colDefs, settings){
  const hdr  = colDefs.map(col=>`<th class="${col.type==='identity'&&col.key==='name'?'col-name-hdr':''}">${escHtml(col.label)}</th>`).join('');
  const rows = results.map(r=>`<tr>${colDefs.map(col=>getCellHTML(r,col,settings)).join('')}</tr>`).join('');
  return`<table class="hitlist-table"><thead><tr>${hdr}</tr></thead><tbody>${rows}</tbody></table>`;
}

function getCellHTML(result, col, settings){
  const {key, type, vaccKey} = col;
  if(type === 'identity'){
    if(key === 'name')    return`<td class="col-name">${escHtml(toTitleCase(result.person.name))}</td>`;
    if(key === 'rank')    return`<td class="col-rank">${escHtml(result.person.rank)}</td>`;
    if(key === 'section') return`<td class="col-section">${escHtml(result.person.section)}</td>`;
  }
  if(type === 'imm-grouped'){
    const grp = result.items.immunizations.grouped;
    if(!grp.includeInReport || grp.count === 0) return`<td class="col-item cell-na">—</td>`;
    const tooltipData = JSON.stringify(grp.dueItems.map(i=>({label:i.label, due:i.due?i.due.getTime():null})));
    return`<td class="col-item ${statusToCss(grp.status)}"><div class="imm-grouped" data-tooltip="${escHtml(tooltipData)}">${grp.count} due</div></td>`;
  }
  if(type === 'imm-individual'){
    const pv = result.items.immunizations.perVaccine[vaccKey];
    if(!pv || !pv.includeInReport) return`<td class="col-item cell-na">—</td>`;
    return`<td class="col-item ${statusToCss(pv.status)}">${pv.due ? formatDate(pv.due) : 'Due'}</td>`;
  }
  if(type === 'tst-combined'){
    const tc = result.items.tstCombined;
    if(!tc || !tc.includeInReport) return`<td class="col-item cell-na">—</td>`;
    const items = tc.tooltipItems || [];
    // Display the most urgent due date
    const displayDue = items.length === 1
      ? items[0].due
      : items.reduce((a,b) => {
          if(!a.due) return b; if(!b.due) return a;
          return a.due <= b.due ? a : b;
        }, items[0])?.due;
    const dateStr = displayDue ? formatDate(displayDue) : 'Due';
    const tooltipData = JSON.stringify(items.map(i=>({label:i.label, due:i.due?i.due.getTime():null})));
    return`<td class="col-item ${statusToCss(tc.status)} tst-combined-cell" data-tooltip="${escHtml(tooltipData)}">${dateStr}</td>`;
  }
  if(type === 'item'){
    const ir = result.items[key];
    if(!ir || !ir.includeInReport) return`<td class="col-item cell-na">—</td>`;
    if(ir.displayText) return`<td class="col-item ${statusToCss(ir.status)}">${escHtml(ir.displayText)}</td>`;
    const due     = getDueDateForItem(result.person, key);
    const dateStr = due ? formatDate(due) : 'Due';
    return`<td class="col-item ${statusToCss(ir.status)}">${dateStr}</td>`;
  }
  return`<td class="col-item cell-na">—</td>`;
}

// Wire fixed-position tooltips for grouped immunization cells after render.
function wireImmTooltips() {
  const tip = dom.floatingTip;
  document.querySelectorAll('.imm-grouped').forEach(cell => {
    let items = [];
    try { items = JSON.parse(cell.dataset.tooltip || '[]'); } catch(e) { return; }
    if(!items.length) return;
    cell.addEventListener('mouseenter', e => {
      const rows = items.map(i => {
        const dateStr = i.due ? formatDateMD(new Date(i.due)) : '';
        return `<div class="tip-imm-row">
          <span class="tip-imm-name">${escHtml(i.label)}</span>
          <span class="tip-imm-date">${escHtml(dateStr)}</span>
        </div>`;
      }).join('');
      tip.innerHTML = `<strong>Due immunizations</strong><div class="tip-imm-grid">${rows}</div>`;
      tip.classList.add('visible');
      placeImmTip(e);
    });
    cell.addEventListener('mousemove', placeImmTip);
    cell.addEventListener('mouseleave', () => tip.classList.remove('visible'));
  });
}

function placeImmTip(e) {
  const tip = dom.floatingTip;
  const PAD = 14;
  let x = e.clientX + PAD, y = e.clientY + PAD;
  tip.style.left = x + 'px'; tip.style.top = y + 'px';
  const r = tip.getBoundingClientRect();
  if(r.right  > window.innerWidth  - 8) x = e.clientX - r.width  - PAD;
  if(r.bottom > window.innerHeight - 8) y = e.clientY - r.height - PAD;
  tip.style.left = x + 'px'; tip.style.top = y + 'px';
}

// Wire fixed-position tooltips for combined TST cells.
// Same pattern as wireImmTooltips — shares placeImmTip and floating-tip element.
function wireTstTooltip() {
  const tip = dom.floatingTip;
  document.querySelectorAll('.tst-combined-cell').forEach(cell => {
    let items = [];
    try { items = JSON.parse(cell.dataset.tooltip || '[]'); } catch(e) { return; }
    if(items.length <= 1) return; // only show tooltip when both are due
    cell.addEventListener('mouseenter', e => {
      const rows = items.map(i => {
        const dateStr = i.due ? formatDateMD(new Date(i.due)) : '';
        return `<div class="tip-imm-row">
          <span class="tip-imm-name">${escHtml(i.label)}</span>
          <span class="tip-imm-date">${escHtml(dateStr)}</span>
        </div>`;
      }).join('');
      tip.innerHTML = `<strong>TST items due</strong><div class="tip-imm-grid">${rows}</div>`;
      tip.classList.add('visible');
      placeImmTip(e);
    });
    cell.addEventListener('mousemove', placeImmTip);
    cell.addEventListener('mouseleave', () => tip.classList.remove('visible'));
  });
}

function getDueDateForItem(person, key){
  const map={
    pha:person.phaDue, dental:person.dentalDue, hiv:person.hivDue, audio:person.audioDue,
    pdha:person.pdhaDue, pdhra:person.pdhraDue,
    mha2:person.mha2?.due||null, mha3:person.mha3?.due||null, mha4:person.mha4?.due||null,
    warningTag:person.warningTagDue, verifyGlasses:person.verifyGlassesDue, verifyInserts:person.verifyInsertsDue,
    refAudio:person.refAudioDue, dna:person.dnaDue, g6pd:person.g6pdDue, sickle:person.sickleDue,
    tst:person.tstDue, tstQuest:person.tstQuestDue,
    'tst-combined': person.tstDue && person.tstQuestDue
      ? (person.tstDue <= person.tstQuestDue ? person.tstDue : person.tstQuestDue)
      : (person.tstDue || person.tstQuestDue),
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

  // TST sub-option — show combine checkbox when either TST item is checked
  function updateTstSub() {
    const eitherChecked = dom.itemTst.checked || dom.itemTstQuest.checked;
    const bothChecked   = dom.itemTst.checked && dom.itemTstQuest.checked;
    dom.tstSub.classList.toggle('hidden', !eitherChecked);
    if(!bothChecked && dom.itemTstCombine) dom.itemTstCombine.checked = false;
    refreshColumnCounter();
  }
  dom.itemTst.addEventListener('change', updateTstSub);
  dom.itemTstQuest.addEventListener('change', updateTstSub);

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
  const validateThresholds = debounce(() => {
    const y = parseInt(dom.threshYellow.value,10), g = parseInt(dom.threshGreen.value,10);
    const invalid = isNaN(y)||isNaN(g)||y<1||g<1||y>=g;
    if(isNaN(y)||isNaN(g)||y<1||g<1) dom.threshError.textContent = 'Thresholds must be positive numbers.';
    else if(y >= g)                   dom.threshError.textContent = 'Yellow window must be smaller than Green window.';
    else                              dom.threshError.textContent = '';
    // Highlight inputs
    dom.threshYellow.classList.toggle('input-error', invalid);
    dom.threshGreen.classList.toggle('input-error',  invalid);
    updateGenerateBtnState();
    refreshColumnCounter();
    saveSettings();
  }, DEBOUNCE_MS);
  dom.threshYellow.addEventListener('input', validateThresholds);
  dom.threshGreen.addEventListener('input',  validateThresholds);

  // Emblem upload
  dom.emblemInput.addEventListener('change',e=>{
    const file=e.target.files[0]; if(!file) return;
    dom.emblemError.classList.add('hidden');
    if(file.size > EMBLEM_MAX_BYTES){
      dom.emblemError.textContent = `Image too large (${formatBytes(file.size)}). Please use a file under 150 KB.`;
      dom.emblemError.classList.remove('hidden');
      dom.emblemInput.value='';
      return;
    }
    const reader=new FileReader();
    reader.onload=ev=>{
      state.emblemBase64=ev.target.result;
      dom.emblemPreview.style.backgroundImage=`url('${ev.target.result}')`;
      dom.emblemPreview.classList.add('has-image'); dom.clearEmblemBtn.classList.remove('hidden');
      onSettingsChanged();
      liveUpdateEmblem();
    };
    reader.readAsDataURL(file);
  });
  dom.clearEmblemBtn.addEventListener('click',()=>{
    state.emblemBase64=null; dom.emblemPreview.style.backgroundImage='';
    dom.emblemPreview.classList.remove('has-image'); dom.clearEmblemBtn.classList.add('hidden');
    dom.emblemInput.value=''; dom.emblemError.classList.add('hidden');
    onSettingsChanged();
    liveUpdateEmblem();
  });

  // Dental date radio
  document.querySelectorAll('input[name="dentalDate"]').forEach(r=>r.addEventListener('change',onSettingsChanged));

  // Text inputs
  dom.unitName.addEventListener('input',debounce(()=>{ onSettingsChanged(); liveUpdateUnitName(); },DEBOUNCE_MS));

  // Info column builder
  initInfoColumns();
  dom.infoColCount.addEventListener('change', () => { initInfoColumns(); onSettingsChanged(); });

  refreshColumnCounter();
}

/*
  Info text multi-column editor.
  Builds N column inputs (textarea + alignment select) based on the
  selected column count.
*/
function initInfoColumns(preload){
  const count=parseInt(dom.infoColCount.value,10)||1;
  // Use preloaded data (from applySettings) or preserve existing editor values
  let existing=[];
  if(Array.isArray(preload)){
    existing=preload;
  } else {
    dom.infoColContainer.querySelectorAll('.info-col-editor').forEach(ed=>{
      existing.push({
        text:  ed.querySelector('textarea').value,
        align: ed.querySelector('select').value,
      });
    });
  }
  dom.infoColContainer.innerHTML='';
  for(let i=0;i<count;i++){
    const prev=existing[i]||{text:'',align:'center'};
    const wrapper=document.createElement('div');
    wrapper.className='info-col-editor';
    wrapper.innerHTML=`
      <div class="info-col-header">
        <span class="info-col-num">Column ${i+1}</span>
        <div class="align-control">
          <span class="align-label">Text alignment:</span>
          <select class="info-align-select align-select">
            <option value="left"${prev.align==='left'?' selected':''}>Left</option>
            <option value="center"${prev.align==='center'?' selected':''}>Center</option>
            <option value="right"${prev.align==='right'?' selected':''}>Right</option>
          </select>
        </div>
      </div>
      <textarea class="settings-textarea info-col-textarea" rows="4"
        placeholder="Clinic hours, phone numbers...">${escHtml(prev.text)}</textarea>`;
    const ta  = wrapper.querySelector('textarea');
    const sel = wrapper.querySelector('select');
    ta.addEventListener('input',   () => { liveUpdateInfoBar(); saveSettings(); });
    sel.addEventListener('change', () => { liveUpdateInfoBar(); saveSettings(); });
    dom.infoColContainer.appendChild(wrapper);
  }
  // onSettingsChanged() is intentionally NOT called here.
  // User-triggered saves are handled by the infoColCount change listener.
  // applySettings() passes preload data and must not trigger a save mid-restore.
}

// Instantly refreshes the info bar in an already-rendered report when column
// text or alignment changes — no re-generate needed.
function liveUpdateInfoBar() {
  if(!state.reportGenerated) return;
  const settings = getSettingsFromUI();
  // Rebuild only the info bar element inside every rendered hitlist-wrapper
  document.querySelectorAll('.hitlist-info-bar, .hitlist-no-info').forEach(el => el.remove());
  document.querySelectorAll('.hitlist-stats-bar').forEach(statsBar => {
    const newBar = buildInfoBar(settings);
    if(newBar) statsBar.insertAdjacentHTML('afterend', newBar);
  });
  // saveSettings() is called by the textarea/select listeners directly so
  // saving is never gated on reportGenerated state.
}

// Instantly refreshes the unit name in the report header without re-generating.
function liveUpdateUnitName() {
  if(!state.reportGenerated) return;
  const name = dom.unitName.value.trim() || 'Unit Name';
  document.querySelectorAll('.hitlist-unit-name').forEach(el => {
    el.textContent = `${name} Medical Hit List`;
  });
}

// Instantly refreshes the emblem in the report header without re-generating.
function liveUpdateEmblem() {
  if(!state.reportGenerated) return;
  const src = state.emblemBase64;
  document.querySelectorAll('.hitlist-header-top').forEach(top => {
    // Replace all emblem elements (img or placeholder divs) within the header-top
    top.querySelectorAll('.hitlist-emblem, .hitlist-emblem-placeholder').forEach(el => {
      if(src) {
        const img = document.createElement('img');
        img.className = 'hitlist-emblem';
        img.src = src;
        img.alt = 'Unit Emblem';
        el.replaceWith(img);
      } else {
        const div = document.createElement('div');
        div.className = 'hitlist-emblem-placeholder';
        div.innerHTML = 'Unit<br/>Emblem';
        el.replaceWith(div);
      }
    });
  });
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

function onSettingsChanged(){ refreshColumnCounter(); saveSettings(); }

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
    infoColCount: parseInt(dom.infoColCount.value, 10) || 1,
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
      tstCombine:    !!(dom.itemTstCombine?.checked && dom.itemTst.checked && dom.itemTstQuest.checked),
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
  const settings = getSettingsFromUI();

  if(state.personnel.length > 0) {
    // File is loaded — run the same evaluation pipeline the report uses
    // so the count reflects real suppression (zero-entry columns excluded)
    const today = new Date(); today.setHours(0,0,0,0);
    const allEval = state.personnel.map(p => evaluatePerson(p, settings, today, settings.projectionDate));
    const colDefs = getColumnDefs(settings);
    const activeDefs = getActiveColumnDefs(colDefs, allEval);
    updateColumnCounter(activeDefs.length, true);
  } else {
    // No file yet — fall back to naive estimate
    updateColumnCounter(getColumnDefs(settings).length, false);
  }
}


/* ── Generate button ──────────────────────────────────────── */

/* ── Generate button state ────────────────────────────────── */

function updateGenerateBtnState() {
  const y = parseInt(dom.threshYellow.value,10);
  const g = parseInt(dom.threshGreen.value,10);
  const threshInvalid = isNaN(y)||isNaN(g)||y<1||g<1||y>=g;
  const fileLoaded    = !!(state.rawData && state.personnel.length);
  const blocked       = threshInvalid; // file gate handled separately via disabled attr on load

  dom.generateBtn.disabled = !fileLoaded || blocked;
  if(dom.generateBlockMsg){
    dom.generateBlockMsg.classList.toggle('hidden', !blocked);
  }
}

dom.generateBtn.addEventListener('click',()=>{
  if(!state.rawData||!state.personnel.length){alert('Please upload a MRRS file first.');return;}

  try {
    const settings = getSettingsFromUI();
    const today    = new Date();
    today.setHours(0,0,0,0);

    // Re-parse if dental date method changed
    if(settings.dentalUseMrrsDate !== state._lastDentalFlag){
      const reparsed = parseMRRS(state.rawData, settings.dentalUseMrrsDate);
      state.personnel = reparsed.personnel; state._lastDentalFlag = settings.dentalUseMrrsDate;
    }

    const {results, stats} = applyFilters(state.personnel, settings, today, settings.projectionDate);
    state.filteredResults = results; state.currentStats = stats;
    renderReport(results, stats, settings, today, settings.projectionDate);
  } catch(err) {
    let context = {};
    try {
      context = {
        personnelCount: state.personnel.length,
        selectedItems: Object.entries(getSettingsFromUI().items).filter(([k,v])=>v).map(([k])=>k),
        immMode: state.immDisplayMode,
        offEnlFilter: dom.offEnlFilter.value,
        columnCount: dom.columnCount.textContent,
      };
    } catch(_) { /* context unavailable */ }
    captureError({
      type: 'generate-error',
      message: err.message,
      stack: err.stack || '',
      context,
    });
    console.error('Report generation error:', err);
    showErrorBar('generate');
  }
});


/* ── 10. SETTINGS PERSISTENCE ──────────────────────────────── */

/* ── getDefaultSettings ─────────────────────────────────────
   Single source of truth for factory defaults.
   applySettings() uses this as the fallback for any key absent
   from a saved or imported settings object.                    */
function getDefaultSettings() {
  return {
    unitName:        '',
    infoColCount:    1,
    infoColumns:     [{ text: '', align: 'center' }],
    emblemBase64:    null,
    items: {
      pha:           true,
      dental:        true,
      hiv:           true,
      audio:         true,
      pdha:          false,
      pdhra:         false,
      mha2:          false,
      mha3:          false,
      mha4:          false,
      warningTag:    false,
      verifyGlasses: false,
      verifyInserts: false,
      bloodType:     true,
      refAudio:      true,
      dna:           true,
      g6pd:          true,
      sickle:        true,
      tst:           false,
      tstQuest:      false,
      tstCombine:    false,
      wellWoman:     false,
      immunizations: true,
    },
    immunizationKeys:  [...IMMUNIZATION_KEYS], // all 20 checked
    immunDisplayMode:  'grouped',
    thresholds:        { yellow: 7, green: 30 },
    offEnlFilter:      'combined',
    sortBy:            'name',
    showRank:          true,
    showSection:       true,
    dentalUseMrrsDate: false,
  };
}

/* ── saveSettings ────────────────────────────────────────────
   Serializes current UI state to localStorage.
   Called automatically by onSettingsChanged().               */
function saveSettings() {
  try {
    const s = getSettingsFromUI();
    // projectionDate is a Date object — store as ISO string
    const payload = Object.assign({}, s, {
      projectionDate: s.projectionDate instanceof Date
        ? s.projectionDate.toISOString()
        : null,
    });
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch(e) {
    console.warn('Hitlist Hatcher: could not save settings', e);
  }
}

/* ── applySettings ───────────────────────────────────────────
   Restores all UI controls from a settings object.
   Missing keys fall back to getDefaultSettings() values.     */
function applySettings(s) {
  const d = getDefaultSettings();
  const g = (key, fallback) => (s[key] !== undefined ? s[key] : fallback);

  // Unit name
  dom.unitName.value = g('unitName', d.unitName);

  // Emblem
  const emb = g('emblemBase64', null);
  if(emb) {
    state.emblemBase64 = emb;
    dom.emblemPreview.style.backgroundImage = `url('${emb}')`;
    dom.emblemPreview.classList.add('has-image');
    dom.clearEmblemBtn.classList.remove('hidden');
  } else {
    state.emblemBase64 = null;
    dom.emblemPreview.style.backgroundImage = '';
    dom.emblemPreview.classList.remove('has-image');
    dom.clearEmblemBtn.classList.add('hidden');
  }

  // Report items checkboxes
  const items = Object.assign({}, d.items, s.items || {});
  dom.itemPha.checked           = !!items.pha;
  dom.itemDental.checked        = !!items.dental;
  dom.itemHiv.checked           = !!items.hiv;
  dom.itemAudio.checked         = !!items.audio;
  dom.itemPdha.checked          = !!items.pdha;
  dom.itemPdhra.checked         = !!items.pdhra;
  dom.itemMha2.checked          = !!items.mha2;
  dom.itemMha3.checked          = !!items.mha3;
  dom.itemMha4.checked          = !!items.mha4;
  dom.itemWarningTag.checked    = !!items.warningTag;
  dom.itemVerifyGlasses.checked = !!items.verifyGlasses;
  dom.itemVerifyInserts.checked = !!items.verifyInserts;
  dom.itemBloodType.checked     = !!items.bloodType;
  dom.itemRefAudio.checked      = !!items.refAudio;
  dom.itemDna.checked           = !!items.dna;
  dom.itemG6pd.checked          = !!items.g6pd;
  dom.itemSickle.checked        = !!items.sickle;
  dom.itemTst.checked           = !!items.tst;
  dom.itemTstQuest.checked      = !!items.tstQuest;
  if(dom.itemTstCombine) dom.itemTstCombine.checked = !!items.tstCombine;
  dom.itemWellWoman.checked     = !!items.wellWoman;

  // Dental sub-option
  dom.dentalSub.classList.toggle('hidden', !items.dental);
  const dentalUseMrrs = g('dentalUseMrrsDate', false);
  const dentalRadios  = document.querySelectorAll('input[name="dentalDate"]');
  dentalRadios.forEach(r => { r.checked = (r.value === 'mrrs') === dentalUseMrrs; });

  // Immunization mode
  const immMode = g('immunDisplayMode', d.immunDisplayMode);
  state.immDisplayMode = immMode;
  dom.immModeGroupedBtn.classList.toggle('active',     immMode === 'grouped');
  dom.immModeIndividualBtn.classList.toggle('active',  immMode === 'individual');

  // Immunization checkboxes
  const savedKeys = new Set(g('immunizationKeys', d.immunizationKeys));
  dom.immCheckboxes.querySelectorAll('.imm-cb').forEach(cb => {
    cb.checked = savedKeys.has(cb.dataset.key);
  });

  // Display options
  dom.offEnlFilter.value = g('offEnlFilter', d.offEnlFilter);
  dom.sortBy.value       = g('sortBy',       d.sortBy);
  dom.showRank.checked    = g('showRank',    d.showRank);
  dom.showSection.checked = g('showSection', d.showSection);

  // Warning thresholds
  const thresh = Object.assign({}, d.thresholds, s.thresholds || {});
  dom.threshYellow.value = thresh.yellow;
  dom.threshGreen.value  = thresh.green;

  // Info columns — delegate to initInfoColumns with preloaded data so
  // column-rebuild logic lives in exactly one place.
  const colCount = g('infoColCount', d.infoColCount);
  dom.infoColCount.value = colCount;
  const savedCols = g('infoColumns', d.infoColumns);
  initInfoColumns(savedCols);

  // TST sub visibility — run through updateTstSub for full consistency
  // (visibility + combine auto-clear when only one box checked)
  dom.itemTst.dispatchEvent(new Event('change'));

  refreshColumnCounter();
}

/* ── loadSettings ────────────────────────────────────────────
   Called once on page load. Reads localStorage and applies.
   First-run (no saved data) applies factory defaults silently. */
function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if(!raw) return; // first run — HTML defaults already in place
    const s = JSON.parse(raw);
    applySettings(s);
  } catch(e) {
    console.warn('Hitlist Hatcher: could not load settings', e);
  }
}

/* ── wireSettingsHandlers ────────────────────────────────────
   Wires Export and Import buttons.                           */
function wireSettingsHandlers() {
  // Export — write JSON content to a .txt file
  dom.exportSettingsBtn.addEventListener('click', () => {
    try {
      const s = getSettingsFromUI();
      const payload = Object.assign({}, s, {
        projectionDate: s.projectionDate instanceof Date
          ? s.projectionDate.toISOString()
          : null,
        _version: APP_VERSION,
        _exported: new Date().toISOString(),
      });
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'text/plain' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = 'hitlist-hatcher-settings.txt';
      a.click();
      URL.revokeObjectURL(url);
    } catch(e) {
      alert('Export failed: ' + e.message);
    }
  });

  // Import — accept .txt or .json, parse as JSON, validate, apply
  dom.importSettingsInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const s = JSON.parse(ev.target.result);
        if(typeof s !== 'object' || s === null) throw new Error('Invalid settings file.');
        applySettings(s);
        saveSettings();
      } catch(err) {
        alert('Could not import settings: ' + err.message + '\nMake sure you are using a settings file exported from Hitlist Hatcher.');
      } finally {
        dom.importSettingsInput.value = ''; // reset so same file can be re-imported
      }
    };
    reader.readAsText(file);
  });
}


/* ── 11. PDF EXPORT ────────────────────────────────────────── */

// A4 landscape at 96 CSS dpi, minus standard margins:
//   width:  (11.69 - 0.35 - 0.35) × 96 = 1055 px
//   height: (8.27  - 0.35 - 0.45) × 96 =  718 px
const PRINT_WIDTH_PX  = 1055;
const PRINT_PAGE_H_PX =  718;

// Core export helper — used by all three export buttons.
//
// Strategy
// ─────────
// position:fixed on .hitlist-header / .hitlist-legend causes Chrome to
// render those elements at the top of every printed page automatically.
//
// The overlap problem (fixed header obscures top rows on pages 2+) is solved
// with two complementary mechanisms:
//   • Page 1 : padding-top on .hitlist-table-wrapper (set via #printHeaderPad)
//              pushes the table below the fixed header.
//   • Pages 2+: a column-name clone <tr> with class .print-col-names-pagebreak
//              is injected before every P-th data row. It carries break-before:page
//              (forces the page break) and padding-top:headerH on each <th>
//              (reserves the fixed-header zone), so column names appear flush
//              below the header with no gap and no separate spacer row.
//
// P (rows per page) is calculated from live measurements at print dimensions
// so the layout is accurate regardless of column count or zoom level.
function exportWrapperAsPdf(wrapperId) {
  const source = wrapperId
    ? document.getElementById(wrapperId)
    : dom.reportOutput.querySelector('.hitlist-wrapper');
  if (!source) return;

  try {
    // Clone the full wrapper (original DOM structure) into printOutput
    dom.printOutput.innerHTML = source.outerHTML;

    // ── 1. Measure heights at print dimensions ───────────────────
    // Temporarily reveal off-screen at the correct print width so that
    // offsetHeight returns accurate values for header and row elements.
    dom.printOutput.style.cssText =
      `display:block;visibility:hidden;position:absolute;` +
      `top:-99999px;left:0;width:${PRINT_WIDTH_PX}px;`;

    // Apply print font-size overrides inline so measured heights match
    // what the browser will actually render in @media print.
    const pName  = dom.printOutput.querySelector('.hitlist-unit-name');
    const pTitle = dom.printOutput.querySelector('.hitlist-title-line');
    if (pName)  pName.style.fontSize  = '18px';
    if (pTitle) pTitle.style.fontSize = '12px';

    const pHeader  = dom.printOutput.querySelector('.hitlist-header');
    const pLegend  = dom.printOutput.querySelector('.hitlist-legend');
    const pHeaderH = pHeader ? pHeader.offsetHeight : 0;   // header block only
    const pLegendH = pLegend ? pLegend.offsetHeight : 0;
    const headerH  = pHeaderH + pLegendH;                  // combined (for padding/spacers)

    // Apply print td/th overrides inline so rowH and colNamesH match actual
    // print rendering (print CSS: font-size:10px, padding:4px 3px).
    const allCells = dom.printOutput.querySelectorAll('.hitlist-table td, .hitlist-table th');
    allCells.forEach(el => {
      el.style.fontSize = '10px';
      el.style.padding  = '4px 3px';
    });

    const colNamesRow = dom.printOutput.querySelector('.hitlist-table thead tr');
    const colNamesH   = colNamesRow ? colNamesRow.offsetHeight : 22;

    const firstRow = dom.printOutput.querySelector('.hitlist-table tbody tr');
    const rowH     = Math.max(firstRow ? firstRow.offsetHeight : 20, 1);

    // Reset measurement state — printOutput returns to display:none via CSS
    dom.printOutput.style.cssText = '';
    if (pName)  pName.style.fontSize  = '';
    if (pTitle) pTitle.style.fontSize = '';
    allCells.forEach(el => {
      el.style.fontSize = '';
      el.style.padding  = '';
    });

    // ── 2. Calculate rows per page ───────────────────────────────
    // Subtract 2 rows as a safety margin against measurement rounding.
    const P = Math.max(5, Math.floor((PRINT_PAGE_H_PX - headerH - colNamesH) / rowH) - 1);

    // ── 3. Insert spacer + column-name clone rows into tbody ───────
    // Before every P-th data row, inject two rows:
    //   a) A spacer <tr> with break-before:page and height:headerH —
    //      this forces a new page and reserves the zone the fixed
    //      header occupies so it doesn't overlap real data.
    //   b) A clone of the <thead> column-name row — gives each page
    //      its own visible column headers below the fixed header zone.
    const tbody      = dom.printOutput.querySelector('.hitlist-table tbody');
    const colNamesEl = dom.printOutput.querySelector('.hitlist-table thead tr');
    const rows       = tbody ? Array.from(tbody.querySelectorAll(':scope > tr')) : [];

    // Determine insert positions: before rows[P], rows[2P], rows[3P]…
    const insertAt = [];
    for (let i = P; i < rows.length; i += P) insertAt.push(i);

    // ── 3a. Insert page-break colClones (pages 2+) ──────────────
    // Strategy: no separate spacer row. Instead, the colClone itself carries
    // break-before:page to force the page break, and padding-top:headerH on
    // each <th> to push the column-name text below the fixed header that
    // overlays the top of every new page. This avoids the height-mismatch gap
    // that a separate spacer row introduced.
    // Insert in reverse order so earlier row indices stay valid.
    for (let j = insertAt.length - 1; j >= 0; j--) {
      if (colNamesEl && tbody) {
        const clone = colNamesEl.cloneNode(true);
        clone.className = 'print-col-names-repeat print-col-names-pagebreak';
        tbody.insertBefore(clone, rows[insertAt[j]]);
      }
    }

    // ── 3b. Insert page-1 colClone at very start of tbody ───────
    // <thead> is suppressed in print (injected CSS below) to stop Chrome from
    // rendering its own unpredictable repetition. This clone replaces it on
    // page 1. The .hitlist-table-wrapper padding-top already clears the header.
    if (colNamesEl && tbody) {
      const firstClone = colNamesEl.cloneNode(true);
      firstClone.className = 'print-col-names-repeat print-col-names-first';
      tbody.insertBefore(firstClone, tbody.firstChild);
    }

    // ── 4. Inject dynamic print styles ──────────────────────────
    let styleEl = document.getElementById('printHeaderPad');
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = 'printHeaderPad';
      document.head.appendChild(styleEl);
    }
    styleEl.textContent =
      `@media print {
         /* Page 1: push table below fixed header */
         .print-only .hitlist-table-wrapper { padding-top: ${headerH}px; }
         /* Legend: position directly below header block */
         .print-only .hitlist-legend { top: ${pHeaderH}px; }
         /* Suppress native <thead> repetition — we own column names via injected rows */
         .print-only .hitlist-table thead { display: none !important; }
         /* All column-name clone rows: visible in print */
         .print-col-names-repeat { display: table-row !important; }
         .print-col-names-repeat th { font-size: 10px !important; padding: 4px 3px !important; }
         /* Page-break clones: force page break, then padding-top pushes text
            below the fixed header — no separate spacer row needed */
         .print-col-names-pagebreak { break-before: page !important; page-break-before: always !important; }
         .print-col-names-pagebreak th { padding-top: ${headerH + 4}px !important; }
       }`;

    window.print();
  } catch(err) {
    captureError({
      type: 'pdf-export-error',
      message: err.message,
      stack: err.stack || '',
      context: {
        wrapperId: wrapperId || 'combined',
        rowCount: source.querySelectorAll('.hitlist-table tbody tr').length,
        colCount: source.querySelectorAll('.hitlist-table thead th').length,
        headerFound: !!source.querySelector('.hitlist-header'),
        legendFound: !!source.querySelector('.hitlist-legend'),
      },
    });
    console.error('PDF export error:', err);
    showErrorBar('export');
  }
}

function wireExportPdfHandler(){
  dom.exportPdfBtn.addEventListener('click',
    () => exportWrapperAsPdf(null));
  dom.exportPdfOfficerBtn.addEventListener('click',
    () => exportWrapperAsPdf('hitlist-officer'));
  dom.exportPdfEnlistedBtn.addEventListener('click',
    () => exportWrapperAsPdf('hitlist-enlisted'));
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

function formatDateMD(date){
  if(!date||!(date instanceof Date)||isNaN(date))return'';
  return`${date.getMonth()+1}/${date.getDate()}`;
}

function formatDateFull(date){
  if(!date||!(date instanceof Date)||isNaN(date))return'';
  const m=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  return`${date.getDate()} ${m[date.getMonth()]} ${date.getFullYear()}`;
}

function daysDiff(a,b){return Math.round((a-b)/86400000);}

function debounce(fn,delay){let t;return(...args)=>{clearTimeout(t);t=setTimeout(()=>fn(...args),delay);};}

function updateColumnCounter(count, exact){
  if(dom.columnLabel) dom.columnLabel.textContent = exact ? 'Columns: ' : 'Estimated columns: ';
  dom.columnCount.textContent=count;
  dom.columnCounter.classList.remove('warn-yellow','warn-red');
  if(count>=COL_WARN_RED)    dom.columnCounter.classList.add('warn-red');
  else if(count>=COL_WARN_YELLOW) dom.columnCounter.classList.add('warn-yellow');
}

// Display-only title case transform for MRRS names.
// MRRS format: "LAST FIRST [MI] [SUFFIX]" — all uppercase.
// Rules:
//   - Suffixes (JR, SR, II, III, IV, V) → keep uppercase
//   - Single-letter tokens (middle initials) → keep uppercase
//   - Mc prefix (3+ chars after Mc) → McDonald, McAteer, etc.
//   - Apostrophes → O'Brien  Hyphens → Smith-Jones
//   - Everything else → first letter up, rest lower
const TITLE_CASE_SUFFIXES = new Set(['JR','SR','II','III','IV','V']);

function toTitleCase(name) {
  if(!name) return '';
  return name.split(' ').map(token => {
    const upper = token.toUpperCase();
    // Preserve suffixes
    if(TITLE_CASE_SUFFIXES.has(upper)) return upper;
    // Preserve single-letter tokens (middle initials)
    if(token.length === 1) return upper;
    // Apply title case to a single word segment (handles hyphens and apostrophes)
    return titleCaseWord(token);
  }).join(' ');
}

function titleCaseWord(word) {
  // Handle hyphens: split, capitalize each part, rejoin
  if(word.includes('-')) return word.split('-').map(titleCaseWord).join('-');
  // Handle apostrophes: O'BRIEN → O'Brien
  const apos = word.indexOf("'");
  if(apos > 0 && apos < word.length - 1) {
    return titleCaseWord(word.slice(0, apos)) + "'" + titleCaseWord(word.slice(apos + 1));
  }
  // Mc prefix: MCDONALD → McDonald (only if 3+ chars follow Mc)
  if(word.length >= 5 && word.slice(0,2).toUpperCase() === 'MC') {
    return 'Mc' + word.charAt(2).toUpperCase() + word.slice(3).toLowerCase();
  }
  // Default: first letter uppercase, rest lowercase
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}
