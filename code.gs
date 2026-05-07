// ============================================================
//  Code.gs — Web App entry point + all server-side API
// ============================================================

/* ── Web App entry ── */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('💰 Budget Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

/* ── HTML include helper ── */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ══════════════════════════════════════════════════════════
   API — Config
══════════════════════════════════════════════════════════ */
function api_getConfig() {
  try {
    const dynCfg = DB.getDynamicConfig();
    return {
      ...dynCfg,
      months      : CFG.MONTHS,
      currentUser : Session.getActiveUser().getEmail() || 'Unknown',
      userRole    : DB.getCurrentUserRole(),
      appName     : CFG.APP_NAME
    };
  } catch(e) {
    // Fallback to static config if DB not ready yet
    const fyLabels = [];
    const cur = new Date().getFullYear(), curM = new Date().getMonth();
    const curFY = curM >= 3 ? cur : cur - 1;
    for (let y = 2024; y <= curFY + 2; y++) fyLabels.push('FY ' + y + '-' + (y+1));
    return {
      companies: CFG.COMPANIES, departments: CFG.DEPARTMENTS, modules: CFG.MODULES,
      projectTypes: CFG.PROJECT_TYPES, mainHeads: CFG.MAIN_HEADS,
      subHeads: CFG.SUB_HEADS, subSubHeads: CFG.SUB_SUB_HEADS, categories: CFG.CATEGORIES,
      months: CFG.MONTHS, fyLabels, roles: CFG.ROLES, statuses: CFG.STATUS,
      currentUser: Session.getActiveUser().getEmail()||'Unknown',
      userRole: 'Finance Admin', appName: CFG.APP_NAME
    };
  }
}

/* ══════════════════════════════════════════════════════════
   API — Budget Entries
══════════════════════════════════════════════════════════ */
function api_getBudgetEntries(company, fy, dept, module_) {
  try {
    return { ok:true, data: DB.getBudgetEntries({company,fy,dept,module_}) };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getActualEntries(company, fy, dept, module_) {
  try {
    return { ok:true, data: DB.getActualEntries({company,fy,dept,module_}) };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_saveBudgetEntry(entryJson, isActual) {
  try {
    const entry = JSON.parse(entryJson);
    // Check lock
    const approval = DB.getApprovalStatus(entry.company, entry.fy, entry.dept, entry.module);
    if (approval.status === CFG.STATUS.LOCKED && !isActual) {
      return {ok:false, msg:'Budget is Locked. Contact Finance Admin.'};
    }
    return DB.saveBudgetEntry(entry, isActual);
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_deleteBudgetEntry(id, isActual) {
  try {
    return DB.deleteBudgetEntry(id, isActual);
  } catch(e) { return {ok:false, msg:e.message}; }
}

/* ══════════════════════════════════════════════════════════
   API — Approval
══════════════════════════════════════════════════════════ */
function api_getApprovalStatus(company, fy, dept, module_) {
  try {
    return { ok:true, data: DB.getApprovalStatus(company, fy, dept, module_) };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_submitForApproval(company, fy, dept, module_) {
  try {
    const result = DB.setApprovalStatus(company, fy, dept, module_, CFG.STATUS.SUBMITTED, '');
    _notifyOnSubmit(company, fy, dept, module_);
    return result;
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_approveBudget(company, fy, dept, module_, remarks) {
  try {
    const result = DB.setApprovalStatus(company, fy, dept, module_, CFG.STATUS.APPROVED, remarks||'');
    _notifyOnApprove(company, fy, dept, module_);
    return result;
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_lockBudget(company, fy, dept, module_) {
  try {
    return DB.setApprovalStatus(company, fy, dept, module_, CFG.STATUS.LOCKED, 'Locked by Finance Admin');
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_requestRevision(company, fy, dept, module_, reason) {
  try {
    const result = DB.setApprovalStatus(company, fy, dept, module_, CFG.STATUS.REVISION, reason);
    _notifyOnRevision(company, fy, dept, module_, reason);
    return result;
  } catch(e) { return {ok:false, msg:e.message}; }
}

/* ══════════════════════════════════════════════════════════
   API — Reports
══════════════════════════════════════════════════════════ */
function api_getMonthlyTotals(company, fy, dept, module_) {
  try {
    return { ok:true, data: DB.getMonthlyTotals(company, fy, dept, module_) };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getSummaryByProjectType(company, fy, dept, module_) {
  try {
    return { ok:true, data: DB.getSummaryByProjectType(company, fy, dept, module_) };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getMultiFYTotals(company, dept, module_) {
  try {
    return { ok:true, data: DB.getMultiFYTotals(company, dept, module_) };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getConsolidated(company, fy) {
  try {
    return { ok:true, data: DB.getConsolidated(company, fy) };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getAuditLog() {
  try {
    return { ok:true, data: DB.getAuditLog(200) };
  } catch(e) { return {ok:false, msg:e.message}; }
}



/**
 * Is function ko select karke ek baar RUN karein.
 * Ye BCOL ke hisaab se headers set kar dega.
 */
function setupBudgetSheetHeaders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CFG.SHEETS.BUDGET; // 'BudgetEntries'
    let sheet = ss.getSheetByName(sheetName);
    
    // Agar sheet nahi hai toh create karein
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // BCOL object se headers ki list nikaalna
    // Keys: ID, COMPANY, FY, DEPT, etc.
    const headers = Object.keys(BCOL); 
    
    // Header row (Row 1) par values set karna
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Styling (Blue background aur White text)
    headerRange.setBackground('#1565C0')
               .setFontColor('white')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    
    // Row 1 ko freeze karna taaki scroll karne par header dikhta rahe
    sheet.setFrozenRows(1);
    
    // Columns ki width auto-adjust karna
    sheet.autoResizeColumns(1, headers.length);

    Logger.log('Success: BudgetEntries sheet headers created!');
    SpreadsheetApp.getUi().alert('Headers have been added successfully!');

  } catch (e) {
    Logger.log('Error: ' + e.toString());
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

/* ══════════════════════════════════════════════════════════
   API — Users / Settings
══════════════════════════════════════════════════════════ */
function api_getUsers() {
  try {
    return { ok:true, data: DB.getUsers() };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_saveUser(userJson) {
  try {
    DB.saveUser(JSON.parse(userJson));
    return { ok:true };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_deleteUser(email) {
  try {
    DB.deleteUser(email);
    return { ok:true };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getSettings() {
  try {
    const keys = ['FINANCE_ADMIN_EMAIL','HR_HOD_EMAIL','SALES_HOD_EMAIL','OTHER_HOD_EMAIL',
                  'EMAIL_NOTIFICATIONS','AUDIT_ENABLED','DEADLINE_DAY'];
    const result = {};
    keys.forEach(k => result[k] = DB.getSetting(k));
    return { ok:true, data:result };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_saveSettings(settingsJson) {
  try {
    const s = JSON.parse(settingsJson);
    Object.entries(s).forEach(([k,v]) => DB.setSetting(k,v));
    return { ok:true };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_exportPDF(company, fy, dept, module_) {
  try {
    const dbId = PropertiesService.getScriptProperties().getProperty('DB_ID');
    if (!dbId) return {ok:false, msg:'Database not initialized.'};
    const ss = SpreadsheetApp.openById(dbId);
    const sh = ss.getSheetByName(CFG.SHEETS.BUDGET);
    const url = 'https://docs.google.com/spreadsheets/d/' + dbId +
      '/export?format=pdf&size=A3&orientation=landscape&fitw=true&gid=' +
      sh.getSheetId() + '&portrait=false';
    const token = ScriptApp.getOAuthToken();
    const resp  = UrlFetchApp.fetch(url, {headers:{'Authorization':'Bearer '+token}});
    const blob  = resp.getBlob().setName(company+'_'+fy+'_'+dept+'_'+module_+'.pdf');
    const file  = DriveApp.createFile(blob);
    DB.audit('PDF_EXPORT','','',[company,fy,dept,module_].join('|'),'');
    return { ok:true, url: file.getUrl(), name: file.getName() };
  } catch(e) { return {ok:false, msg:e.message}; }
}


/* ══════════════════════════════════════════════════════════
   API — Master Data (Finance Admin only)
══════════════════════════════════════════════════════════ */
function api_getMasterData() {
  try {
    return { ok:true, data: DB.getMasterData() };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_addMasterItem(type, value, parentValue) {
  try {
    if (DB.getCurrentUserRole() !== 'Finance Admin') return {ok:false, msg:'Only Finance Admin can add master data.'};
    return DB.addMasterItem(type, value, parentValue);
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_updateMasterItem(type, oldValue, newValue, parentValue, status) {
  try {
    if (DB.getCurrentUserRole() !== 'Finance Admin') return {ok:false, msg:'Only Finance Admin can update master data.'};
    return DB.updateMasterItem(type, oldValue, newValue, parentValue, status);
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_deleteMasterItem(type, value, parentValue) {
  try {
    if (DB.getCurrentUserRole() !== 'Finance Admin') return {ok:false, msg:'Only Finance Admin can delete master data.'};
    return DB.deleteMasterItem(type, value, parentValue);
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getDynamicConfig() {
  try {
    return { ok:true, data: DB.getDynamicConfig() };
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_resetAndReinitDB() {
  try {
    DB.resetDBCache();
    DB.getDB(); // re-init sheets
    return { ok:true, msg:'DB re-initialized. Refresh the page.' };
  } catch(e) { return {ok:false, msg:e.message}; }
}


/* ══════════════════════════════════════════════════════════
   API — Batch (single call for dashboard — faster)
══════════════════════════════════════════════════════════ */
function api_getDashboardData(company, fy, dept, module_) {
  try {
    return {
      ok          : true,
      monthly     : DB.getMonthlyTotals(company, fy, dept, module_),
      byProjectType: DB.getSummaryByProjectType(company, fy, dept, module_),
      approval    : DB.getApprovalStatus(company, fy, dept, module_),
      consolidated: DB.getConsolidated(company, fy)
    };
  } catch(e) { return {ok:false, msg:e.message}; }
}

/* Batch entry load — budget + actual together */
function api_getAllEntries(company, fy, dept, module_) {
  try {
    return {
      ok     : true,
      budget : DB.getBudgetEntries({company,fy,dept,module_}),
      actual : DB.getActualEntries({company,fy,dept,module_}),
      approval: DB.getApprovalStatus(company, fy, dept, module_)
    };
  } catch(e) { return {ok:false, msg:e.message}; }
}


/* ══════════════════════════════════════════════════════════
   API — Import / Export
══════════════════════════════════════════════════════════ */
function api_exportCSV(company, fy, dept, module_, isActual) {
  try {
    return DB.exportBudgetCSV(company, fy, dept, module_, isActual);
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_importCSV(rowsJson, isActual) {
  try {
    return DB.importBudgetCSV(rowsJson, isActual);
  } catch(e) { return {ok:false, msg:e.message}; }
}

function api_getImportTemplate() {
  const months = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];
  const headers = ['Company','FY','Department','Module','ProjectType','MainHead',
    'SubHead','SubSubHead','Category','SiteName','Recurring','Notes',
    ...months.map(m=>m+'_Budget'),'Status'];
  return { ok:true, headers };
}

function api_resetDBCache() {
  try { DB.resetDBCache(); return {ok:true}; } catch(e) { return {ok:false,msg:e.message}; }
}


function api_unlockBudget(company, fy, dept, module_) {
  try {
    if (DB.getCurrentUserRole() !== 'Finance Admin') return {ok:false, msg:'Only Finance Admin can unlock.'};
    return DB.setApprovalStatus(company, fy, dept, module_, CFG.STATUS.APPROVED, 'Unlocked by Finance Admin');
  } catch(e) { return {ok:false, msg:e.message}; }
}

/* ── Email helpers ── */
function _notifyOnSubmit(company, fy, dept, module_) {
  try {
    const adminEmail = DB.getSetting('FINANCE_ADMIN_EMAIL');
    if (!adminEmail || DB.getSetting('EMAIL_NOTIFICATIONS')==='false') return;
    const user = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: adminEmail,
      subject: '📤 Budget Submitted: '+company+' | '+fy+' | '+dept,
      htmlBody: _emailTemplate('Budget Submitted for Approval',
        [{l:'Company',v:company},{l:'FY',v:fy},{l:'Department',v:dept},
         {l:'Module',v:module_},{l:'Submitted By',v:user},{l:'Action',v:'Please review and Approve'}],
        '','#1565C0')
    });
  } catch(e) { Logger.log('Email error: '+e.message); }
}

function _notifyOnApprove(company, fy, dept, module_) {
  try {
    const hodKey = dept.toUpperCase()+'_HOD_EMAIL';
    const hodEmail = DB.getSetting(hodKey);
    if (!hodEmail || DB.getSetting('EMAIL_NOTIFICATIONS')==='false') return;
    MailApp.sendEmail({
      to: hodEmail,
      subject: '✅ Budget Approved: '+company+' | '+fy+' | '+dept,
      htmlBody: _emailTemplate('Budget Approved & Locked',
        [{l:'Company',v:company},{l:'FY',v:fy},{l:'Department',v:dept},
         {l:'Module',v:module_},{l:'Note',v:'Budget is now APPROVED.'}],
        '','#2E7D32')
    });
  } catch(e) { Logger.log('Email error: '+e.message); }
}

function _notifyOnRevision(company, fy, dept, module_, reason) {
  try {
    const hodKey = dept.toUpperCase()+'_HOD_EMAIL';
    const hodEmail = DB.getSetting(hodKey);
    if (!hodEmail || DB.getSetting('EMAIL_NOTIFICATIONS')==='false') return;
    MailApp.sendEmail({
      to: hodEmail,
      subject: '🔄 Revision Required: '+company+' | '+fy+' | '+dept,
      htmlBody: _emailTemplate('Revision Requested',
        [{l:'Company',v:company},{l:'FY',v:fy},{l:'Department',v:dept},
         {l:'Reason',v:reason},{l:'Action',v:'Please update and re-submit.'}],
        '','#E65100')
    });
  } catch(e) { Logger.log('Email error: '+e.message); }
}

function _emailTemplate(title, rows, url, color) {
  const rowsHtml = rows.map(r=>`<tr><td style="padding:8px;font-weight:bold;background:#F5F5F5;border:1px solid #E0E0E0;width:160px">${r.l}</td><td style="padding:8px;border:1px solid #E0E0E0">${r.v}</td></tr>`).join('');
  return `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto">
    <div style="background:${color};padding:20px;text-align:center"><h2 style="color:#FFF;margin:0">💰 Budget System</h2><p style="color:rgba(255,255,255,.8);margin:6px 0 0">${title}</p></div>
    <div style="padding:20px"><table style="width:100%;border-collapse:collapse">${rowsHtml}</table></div>
    <div style="padding:10px;text-align:center;font-size:11px;color:#999">Auto-generated — do not reply</div></div>`;
}