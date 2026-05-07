// ============================================================
//  Database.gs — All CRUD on hidden Google Sheet (DB)
// ============================================================
const DB = (() => {

  /* ── Get the bound spreadsheet — cached after first init ── */
  function getDB() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const initialized = PropertiesService.getScriptProperties().getProperty('DB_INITIALIZED');
    if (!initialized) {
      _ensureAllSheets(ss);
      PropertiesService.getScriptProperties().setProperty('DB_INITIALIZED', '1');
    }
    return ss;
  }

  /* ── Force re-init (call if new sheet needed) ── */
  function resetDBCache() {
    PropertiesService.getScriptProperties().deleteProperty('DB_INITIALIZED');
    PropertiesService.getScriptProperties().deleteProperty('MASTER_CACHE');
  }

  /* ── Auto-create any missing DB sheets in this workbook ── */
  function _ensureAllSheets(ss) {
    const needed = [
      { name: CFG.SHEETS.BUDGET,   fn: _setupBudgetSheet   },
      { name: CFG.SHEETS.ACTUAL,   fn: _setupBudgetSheet   },
      { name: CFG.SHEETS.APPROVAL, fn: _setupApprovalSheet },
      { name: CFG.SHEETS.USERS,    fn: _setupUsersSheet    },
      { name: CFG.SHEETS.AUDIT,    fn: _setupAuditSheet    },
      { name: CFG.SHEETS.SETTINGS, fn: _setupSettingsSheet },
      { name: CFG.SHEETS.MASTER,   fn: _setupMasterSheet   }
    ];
    needed.forEach(({ name, fn }) => {
      if (!ss.getSheetByName(name)) {
        fn(ss.insertSheet(name));
      }
    });
  }

  function _setupBudgetSheet(sh) {
    const hdrs = ['ID','Company','FY','Department','Module','ProjectType','MainHead','SubHead',
      'SubSubHead','Category','SiteName','Recurring','Notes',
      'Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar',
      'Total','Status','CreatedBy','CreatedOn','ModifiedBy','ModifiedOn'];
    sh.getRange(1,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFF');
    sh.setFrozenRows(1);
  }

  function _setupApprovalSheet(sh) {
    const hdrs = ['ID','Company','FY','Department','Module','Status','Version',
                  'SubmittedBy','SubmittedOn','ApprovedBy','ApprovedOn','Remarks','UpdatedOn'];
    sh.getRange(1,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFF');
    sh.setFrozenRows(1);
  }

  function _setupUsersSheet(sh) {
    const hdrs = ['Email','Name','Role','Company','Department','Status','AddedOn'];
    sh.getRange(1,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFF');
    sh.setFrozenRows(1);
    // Add current user as Finance Admin
    sh.appendRow([Session.getActiveUser().getEmail(),'Administrator','Finance Admin',
                  CFG.COMPANIES[0],'All','Active',new Date().toISOString()]);
  }

  function _setupMasterSheet(sh) {
    const hdrs = ['Type','Value','ParentValue','SortOrder','Status','CreatedBy','CreatedOn'];
    sh.getRange(1,1,1,hdrs.length).setValues([hdrs])
      .setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFF');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 200);
    // Seed default master data from Config
    const defaults = [];
    CFG.COMPANIES.forEach((v,i)       => defaults.push(['Company',     v, '',    i+1, 'Active','System','']));
    CFG.DEPARTMENTS.forEach((v,i)     => defaults.push(['Department',  v, '',    i+1, 'Active','System','']));
    CFG.MODULES.forEach((v,i)         => defaults.push(['Module',      v, '',    i+1, 'Active','System','']));
    CFG.PROJECT_TYPES.forEach((v,i)   => defaults.push(['ProjectType', v, '',    i+1, 'Active','System','']));
    CFG.MAIN_HEADS.forEach((v,i)      => defaults.push(['MainHead',    v, '',    i+1, 'Active','System','']));
    // SubHead: parent = Department (each dept gets all sub heads by default)
    // HR and Sales get Expense sub heads; Other gets both
    const subHeadsByDept = {
      'HR'   : ['Direct Expense','Indirect Expense'],
      'Sales': ['Direct Expense','Indirect Expense','Direct Income','Indirect Income'],
      'Other': ['Direct Expense','Indirect Expense','Direct Income','Indirect Income']
    };
    Object.entries(subHeadsByDept).forEach(([dept, list]) =>
      list.forEach((v,i) => defaults.push(['SubHead', v, dept, i+1, 'Active','System',''])));

    // SubSubHead: parent = SubHead (unchanged)
    Object.entries(CFG.SUB_SUB_HEADS).forEach(([sh_, list]) =>
      list.forEach((v,i) => defaults.push(['SubSubHead', v, sh_, i+1, 'Active','System',''])));

    // Category: parent = SubSubHead (unchanged)
    Object.entries(CFG.CATEGORIES).forEach(([ssh, list]) =>
      list.forEach((v,i) => defaults.push(['Category', v, ssh, i+1, 'Active','System',''])));
    // Default site names
    ['Main Office Site','Site A','Site B','Site C'].forEach((v,i) =>
      defaults.push(['SiteName', v, '', i+1, 'Active','System','']));

    if (defaults.length > 0) sh.getRange(2,1,defaults.length,7).setValues(defaults);
  }

  function _setupAuditSheet(sh) {
    const hdrs = ['Timestamp','User','Action','Entity','OldValue','NewValue','Details'];
    sh.getRange(1,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFF');
    sh.setFrozenRows(1);
  }

  function _setupSettingsSheet(sh) {
    sh.getRange(1,1,1,2).setValues([['Key','Value']]).setFontWeight('bold');
    const defaults = [
      ['FINANCE_ADMIN_EMAIL',''],['HR_HOD_EMAIL',''],['SALES_HOD_EMAIL',''],
      ['OTHER_HOD_EMAIL',''],['EMAIL_NOTIFICATIONS','true'],['AUDIT_ENABLED','true'],
      ['DEADLINE_DAY','15'],['CURRENCY','INR']
    ];
    defaults.forEach(r => sh.appendRow(r));
    sh.setFrozenRows(1);
  }

  /* ── Generic helpers ── */
  function getSheet(name) { return getDB().getSheetByName(name); }

  function getAllRows(sheetName) {
    const sh   = getSheet(sheetName);
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  }

  function appendRow(sheetName, rowArr) {
    getSheet(sheetName).appendRow(rowArr);
  }

  function updateRow(sheetName, rowIndex, rowArr) {
    // rowIndex is 1-based data row (not sheet row — add 1 for header)
    const sheetRow = rowIndex + 1;
    getSheet(sheetName).getRange(sheetRow, 1, 1, rowArr.length).setValues([rowArr]);
  }

  function deleteRow(sheetName, rowIndex) {
    getSheet(sheetName).deleteRow(rowIndex + 1);
  }

  /* ── Settings ── */
  function getSetting(key) {
    const rows = getAllRows(CFG.SHEETS.SETTINGS);
    const row  = rows.find(r => r[0] === key);
    return row ? row[1] : '';
  }

  function setSetting(key, value) {
    const sh   = getSheet(CFG.SHEETS.SETTINGS);
    const rows = getAllRows(CFG.SHEETS.SETTINGS);
    const idx  = rows.findIndex(r => r[0] === key);
    if (idx >= 0) {
      sh.getRange(idx + 2, 2).setValue(value);
    } else {
      sh.appendRow([key, value]);
    }
  }

  /* ── Audit Log ── */
  function audit(action, entity, oldVal, newVal, details) {
    try {
      appendRow(CFG.SHEETS.AUDIT, [
        new Date().toISOString(),
        Session.getActiveUser().getEmail() || 'System',
        action, entity,
        String(oldVal||'').substring(0,300),
        String(newVal||'').substring(0,300),
        String(details||'').substring(0,500)
      ]);
    } catch(e) { Logger.log('Audit error: '+e.message); }
  }

  /* ── Users ── */
  function getUsers() {
    return getAllRows(CFG.SHEETS.USERS).map(r => ({
      email:r[0], name:r[1], role:r[2], company:r[3], dept:r[4], status:r[5], addedOn:r[6]
    }));
  }

  function saveUser(user) {
    const rows = getAllRows(CFG.SHEETS.USERS);
    const idx  = rows.findIndex(r => r[0].toLowerCase() === user.email.toLowerCase());
    const rowArr = [user.email, user.name, user.role, user.company, user.dept, user.status, user.addedOn || new Date().toISOString()];
    if (idx >= 0) updateRow(CFG.SHEETS.USERS, idx, rowArr);
    else appendRow(CFG.SHEETS.USERS, rowArr);
  }

  function deleteUser(email) {
    const rows = getAllRows(CFG.SHEETS.USERS);
    const idx  = rows.findIndex(r => r[0].toLowerCase() === email.toLowerCase());
    if (idx >= 0) deleteRow(CFG.SHEETS.USERS, idx);
  }

  function getCurrentUserRole() {
    const email = (Session.getActiveUser().getEmail()||'').toLowerCase();
    const users = getUsers();
    const u     = users.find(u => u.email.toLowerCase() === email && u.status === 'Active');
    return u ? u.role : 'Viewer';
  }

  /* ── Budget Entries ── */
  function getBudgetEntries(filters) {
    const rows = getAllRows(CFG.SHEETS.BUDGET);
    return rows.filter(r => _matchFilters(r, filters, BCOL)).map(_rowToBudget);
  }

  function getActualEntries(filters) {
    const rows = getAllRows(CFG.SHEETS.ACTUAL);
    return rows.filter(r => _matchFilters(r, filters, ACOL)).map(_rowToBudget);
  }

  function saveBudgetEntry(entry, isActual) {
    const sheetName = isActual ? CFG.SHEETS.ACTUAL : CFG.SHEETS.BUDGET;
    const rows      = getAllRows(sheetName);
    const col       = isActual ? ACOL : BCOL;
    const now       = new Date().toISOString();
    const user      = Session.getActiveUser().getEmail() || 'System';

    const months    = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];
    const monthVals = months.map(m => Number(entry[m])||0);
    const total     = monthVals.reduce((a,v)=>a+v,0);

    if (entry.id) {
      const idx = rows.findIndex(r => r[col.ID] === entry.id);
      if (idx >= 0) {
        const rowArr = _buildRowArr(entry, monthVals, total, rows[idx][col.CREATED_BY], rows[idx][col.CREATED_ON], user, now);
        updateRow(sheetName, idx, rowArr);
        audit('UPDATE', sheetName, '', entry.id, JSON.stringify(entry));
        return { ok:true, id:entry.id };
      }
    }

    const id = 'B' + Date.now();
    const rowArr = _buildRowArr({...entry,id}, monthVals, total, user, now, user, now);
    appendRow(sheetName, rowArr);
    audit('INSERT', sheetName, '', id, JSON.stringify(entry));
    return { ok:true, id };
  }

  function deleteBudgetEntry(id, isActual) {
    const sheetName = isActual ? CFG.SHEETS.ACTUAL : CFG.SHEETS.BUDGET;
    const col       = isActual ? ACOL : BCOL;
    const rows      = getAllRows(sheetName);
    const idx       = rows.findIndex(r => r[col.ID] === id);
    if (idx >= 0) { deleteRow(sheetName, idx); audit('DELETE', sheetName, id, '', ''); }
    return { ok: idx >= 0 };
  }

  /* ── Approval ── */
  function getApprovalStatus(company, fy, dept, module_) {
    const key  = [company, fy, dept, module_].join('|');
    const rows = getAllRows(CFG.SHEETS.APPROVAL);
    const row  = rows.find(r => [r[1],r[2],r[3],r[4]].join('|') === key);
    if (!row) return { status: CFG.STATUS.DRAFT, version: 0, remarks: '' };
    return { id:row[0], status:row[5], version:row[6], submittedBy:row[7], submittedOn:row[8],
             approvedBy:row[9], approvedOn:row[10], remarks:row[11] };
  }

  function setApprovalStatus(company, fy, dept, module_, status, remarks) {
    const key  = [company, fy, dept, module_].join('|');
    const rows = getAllRows(CFG.SHEETS.APPROVAL);
    const now  = new Date().toISOString();
    const user = Session.getActiveUser().getEmail() || 'System';
    const idx  = rows.findIndex(r => [r[1],r[2],r[3],r[4]].join('|') === key);

    let rowArr;
    if (idx >= 0) {
      const cur = rows[idx];
      const version = (Number(cur[6])||0) + 1;
      const submittedBy  = status === CFG.STATUS.SUBMITTED ? user : cur[7];
      const submittedOn  = status === CFG.STATUS.SUBMITTED ? now  : cur[8];
      const approvedBy   = status === CFG.STATUS.APPROVED  ? user : cur[9];
      const approvedOn   = status === CFG.STATUS.APPROVED  ? now  : cur[10];
      rowArr = [cur[0], company, fy, dept, module_, status, version,
                submittedBy, submittedOn, approvedBy, approvedOn, remarks||'', now];
      updateRow(CFG.SHEETS.APPROVAL, idx, rowArr);
    } else {
      const id = 'A' + Date.now();
      rowArr = [id, company, fy, dept, module_, status, 1,
                status===CFG.STATUS.SUBMITTED?user:'', status===CFG.STATUS.SUBMITTED?now:'',
                status===CFG.STATUS.APPROVED?user:'',  status===CFG.STATUS.APPROVED?now:'',
                remarks||'', now];
      appendRow(CFG.SHEETS.APPROVAL, rowArr);
    }
    audit('APPROVAL_'+status, CFG.SHEETS.APPROVAL, '', key, remarks||'');
    return { ok:true };
  }

  /* ── Reports ── */
  function getMonthlyTotals(company, fy, dept, module_) {
    const entries = getBudgetEntries({company, fy, dept, module_});
    const actuals = getActualEntries({company, fy, dept, module_});
    const months  = CFG.MONTHS;
    const result  = months.map(m => ({
      month: m,
      budget: entries.reduce((a,e) => a+(Number(e[m])||0), 0),
      actual: actuals.reduce((a,e) => a+(Number(e[m])||0), 0)
    }));
    result.forEach(r => r.variance = r.budget - r.actual);
    return result;
  }

  function getSummaryByProjectType(company, fy, dept, module_) {
    const entries = getBudgetEntries({company, fy, dept, module_});
    const actuals = getActualEntries({company, fy, dept, module_});
    const map = {};
    entries.forEach(e => {
      const k = e.projectType;
      if (!map[k]) map[k] = {projectType:k, budget:0, actual:0};
      map[k].budget += e.total||0;
    });
    actuals.forEach(e => {
      const k = e.projectType;
      if (!map[k]) map[k] = {projectType:k, budget:0, actual:0};
      map[k].actual += e.total||0;
    });
    return Object.values(map).map(r => ({...r, variance: r.budget - r.actual}));
  }

  function getMultiFYTotals(company, dept, module_) {
    const allBudget = getAllRows(CFG.SHEETS.BUDGET).map(_rowToBudget)
      .filter(e => e.company===company && e.dept===dept && e.module===module_);
    const allActual = getAllRows(CFG.SHEETS.ACTUAL).map(_rowToBudget)
      .filter(e => e.company===company && e.dept===dept && e.module===module_);
    const fyMap = {};
    allBudget.forEach(e => {
      if(!fyMap[e.fy]) fyMap[e.fy]={fy:e.fy,budget:0,actual:0};
      fyMap[e.fy].budget += e.total||0;
    });
    allActual.forEach(e => {
      if(!fyMap[e.fy]) fyMap[e.fy]={fy:e.fy,budget:0,actual:0};
      fyMap[e.fy].actual += e.total||0;
    });
    return Object.values(fyMap).sort((a,b)=>a.fy.localeCompare(b.fy))
      .map(r=>({...r, variance:r.budget-r.actual}));
  }

  function getConsolidated(company, fy) {
    const result = [];
    CFG.DEPARTMENTS.forEach(dept => {
      CFG.MODULES.forEach(module_ => {
        const entries = getBudgetEntries({company,fy,dept,module_});
        const actuals = getActualEntries({company,fy,dept,module_});
        const budget  = entries.reduce((a,e)=>a+(e.total||0),0);
        const actual  = actuals.reduce((a,e)=>a+(e.total||0),0);
        const approval = getApprovalStatus(company,fy,dept,module_);
        result.push({dept,module_,budget,actual,variance:budget-actual,
                     status:approval.status, entries:entries.length});
      });
    });
    return result;
  }

  function getAuditLog(limit) {
    const rows = getAllRows(CFG.SHEETS.AUDIT);
    return rows.slice(-(limit||100)).reverse().map(r=>({
      timestamp:r[0],user:r[1],action:r[2],entity:r[3],
      oldValue:r[4],newValue:r[5],details:r[6]
    }));
  }

  /* ── Private helpers ── */
  function _matchFilters(row, filters, col) {
    if (!filters) return true;
    return Object.entries(filters).every(([k,v]) => {
      if (!v) return true;
      const colIdx = {company:col.COMPANY,fy:col.FY,dept:col.DEPT,
                      module:col.MODULE,projectType:col.PROJECT_TYPE}[k];
      return colIdx !== undefined ? row[colIdx] === v : true;
    });
  }

  function _rowToBudget(r) {
    const months = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'];
    const obj = {
      id:r[BCOL.ID], company:r[BCOL.COMPANY], fy:r[BCOL.FY], dept:r[BCOL.DEPT],
      module:r[BCOL.MODULE], projectType:r[BCOL.PROJECT_TYPE], mainHead:r[BCOL.MAIN_HEAD],
      subHead:r[BCOL.SUB_HEAD], subSubHead:r[BCOL.SUB_SUB_HEAD], category:r[BCOL.CATEGORY],
      siteName:r[BCOL.SITE_NAME], recurring:r[BCOL.RECURRING], notes:r[BCOL.NOTES],
      total:r[BCOL.TOTAL], status:r[BCOL.STATUS], createdBy:r[BCOL.CREATED_BY], modifiedOn:r[BCOL.MODIFIED_ON]
    };
    months.forEach((m,i) => { obj[m] = r[BCOL.APR + i]; });
    return obj;
  }

  function _buildRowArr(e, monthVals, total, createdBy, createdOn, modifiedBy, modifiedOn) {
    return [
      e.id, e.company, e.fy, e.dept, e.module,
      e.projectType, e.mainHead, e.subHead, e.subSubHead, e.category,
      e.siteName||'', e.recurring?'Yes':'No', e.notes||'',
      ...monthVals, total,
      e.status||'Active', createdBy, createdOn, modifiedBy, modifiedOn
    ];
  }

  /* ══ MASTER DATA CRUD ══════════════════════════════════════ */

  function getMasterData() {
    const rows = getAllRows(CFG.SHEETS.MASTER);
    return rows.map(r => ({
      type: r[0], value: r[1], parentValue: r[2],
      sortOrder: r[3], status: r[4], createdBy: r[5], createdOn: r[6]
    }));
  }

  function getMasterByType(type) {
    return getMasterData().filter(r => r.type === type && r.status === 'Active');
  }

  function addMasterItem(type, value, parentValue) {
    const rows  = getAllRows(CFG.SHEETS.MASTER);
    // Check duplicate
    const dup = rows.find(r => r[0]===type && r[1].toLowerCase()===value.toLowerCase() && (r[2]||'')===(parentValue||''));
    if (dup) return { ok:false, msg: value + ' already exists in ' + type };
    const sortOrder = rows.filter(r=>r[0]===type).length + 1;
    const user = Session.getActiveUser().getEmail()||'System';
    getSheet(CFG.SHEETS.MASTER).appendRow([type, value, parentValue||'', sortOrder, 'Active', user, new Date().toISOString()]);
    audit('MASTER_ADD', type, '', value, parentValue||'');
    PropertiesService.getScriptProperties().deleteProperty('MASTER_CACHE');
    return { ok:true };
  }

  function updateMasterItem(type, oldValue, newValue, parentValue, status) {
    const sh   = getSheet(CFG.SHEETS.MASTER);
    const rows = getAllRows(CFG.SHEETS.MASTER);
    const idx  = rows.findIndex(r => r[0]===type && r[1]===oldValue && (r[2]||'')===(parentValue||''));
    if (idx < 0) return { ok:false, msg: 'Item not found' };
    sh.getRange(idx+2, 2).setValue(newValue||oldValue);
    sh.getRange(idx+2, 5).setValue(status||rows[idx][4]);
    audit('MASTER_UPDATE', type, oldValue, newValue||oldValue, '');
    PropertiesService.getScriptProperties().deleteProperty('MASTER_CACHE');
    return { ok:true };
  }

  function deleteMasterItem(type, value, parentValue) {
    const sh   = getSheet(CFG.SHEETS.MASTER);
    const rows = getAllRows(CFG.SHEETS.MASTER);
    const idx  = rows.findIndex(r => r[0]===type && r[1]===value && (r[2]||'')===(parentValue||''));
    if (idx < 0) return { ok:false, msg: 'Item not found' };
    sh.getRange(idx+2, 5).setValue('Inactive'); // soft delete
    audit('MASTER_DELETE', type, value, 'Inactive', '');
    PropertiesService.getScriptProperties().deleteProperty('MASTER_CACHE');
    return { ok:true };
  }


  /* ══ IMPORT / EXPORT ════════════════════════════════════════ */

  function exportBudgetCSV(company, fy, dept, module_, isActual) {
    const entries = isActual ? getActualEntries({company,fy,dept,module_})
                             : getBudgetEntries({company,fy,dept,module_});
    const months = CFG.MONTHS;
    const headers = ['Company','FY','Department','Module','ProjectType','MainHead',
      'SubHead','SubSubHead','Category','SiteName','Recurring','Notes',
      ...months.map(m=>m+'_Budget'), 'Total','Status'];
    const rows = entries.map(e => [
      e.company, e.fy, e.dept, e.module, e.projectType, e.mainHead,
      e.subHead, e.subSubHead, e.category, e.siteName||'',
      e.recurring||'', e.notes||'',
      ...months.map(m => e[m]||0), e.total||0, e.status||''
    ]);
    return { ok:true, headers, rows, filename: [company,fy,dept,module_].join('_')+'.csv' };
  }

  function importBudgetCSV(rowsJson, isActual) {
    try {
      const rows   = JSON.parse(rowsJson);
      const months = CFG.MONTHS;
      let imported = 0, skipped = 0;
      rows.forEach(r => {
        if (!r.Company || !r.FY || !r.Department || !r.Module || !r.Category) { skipped++; return; }
        const entry = {
          company: r.Company, fy: r.FY, dept: r.Department, module: r.Module,
          projectType: r.ProjectType||'', mainHead: r.MainHead||'',
          subHead: r.SubHead||'', subSubHead: r.SubSubHead||'',
          category: r.Category, siteName: r.SiteName||'',
          recurring: r.Recurring==='Yes', notes: r.Notes||'', status: r.Status||'Active'
        };
        months.forEach(m => { entry[m] = Number(r[m+'_Budget'])||0; });
        saveBudgetEntry(entry, isActual);
        imported++;
      });
      audit('IMPORT', isActual?'Actual':'Budget', '', imported+' rows', '');
      return { ok:true, imported, skipped };
    } catch(e) { return {ok:false, msg:e.message}; }
  }

  /* Get all master data grouped as config object (replaces CFG for frontend) */
  function getDynamicConfig(forceRefresh) {
    if (!forceRefresh) {
      const cached = PropertiesService.getScriptProperties().getProperty('MASTER_CACHE');
      if (cached) {
        try { return JSON.parse(cached); } catch(e) {}
      }
    }
    const all = getMasterData().filter(r => r.status === 'Active');
    const byType = {};
    all.forEach(r => {
      if (!byType[r.type]) byType[r.type] = [];
      byType[r.type].push({ value: r.value, parent: r.parentValue });
    });
    const vals = t => {
  const list = (byType[t] || []).map(r => r.value);
   return [...new Set(list)]; // 🔥 removes duplicates
    };
    // subHeads keyed by Department  e.g. { HR: ['Direct Expense', ...], Sales: [...] }
    // subSubHeads keyed by SubHead   e.g. { 'Direct Expense': ['Labour Cost', ...] }
    // categories keyed by SubSubHead e.g. { 'Labour Cost': ['Wages', ...] }
    const subHeads = {}, subSubHeads = {}, categories = {};
    (byType['SubHead']||[]).forEach(r => {
      if (!subHeads[r.parent]) subHeads[r.parent] = [];
      if (!subHeads[r.parent].includes(r.value)) subHeads[r.parent].push(r.value);
    });
    (byType['SubSubHead']||[]).forEach(r => {
      if (!subSubHeads[r.parent]) subSubHeads[r.parent] = [];
      if (!subSubHeads[r.parent].includes(r.value)) subSubHeads[r.parent].push(r.value);
    });
    (byType['Category']||[]).forEach(r => {
      if (!categories[r.parent]) categories[r.parent] = [];
      if (!categories[r.parent].includes(r.value)) categories[r.parent].push(r.value);
    });
    const fyLabels = [];
    const cur = new Date().getFullYear(), curM = new Date().getMonth();
    const curFY = curM >= 3 ? cur : cur-1;
    for (let y=2024; y<=curFY+2; y++) fyLabels.push('FY '+y+'-'+(y+1));
    const result = {
      companies   : vals('Company'),
      departments : vals('Department'),
      modules     : vals('Module'),
      projectTypes: vals('ProjectType'),
      mainHeads   : vals('MainHead'),
      subHeads, subSubHeads, categories,
      siteNames   : vals('SiteName'),
      fyLabels,
      roles       : CFG.ROLES,
      statuses    : CFG.STATUS
    };
    try { PropertiesService.getScriptProperties().setProperty('MASTER_CACHE', JSON.stringify(result)); } catch(e) {}
    return result;
  }

  return {
    getDB, getSheet, getUsers, saveUser, deleteUser, getCurrentUserRole,
    getBudgetEntries, getActualEntries, saveBudgetEntry, deleteBudgetEntry,
    getApprovalStatus, setApprovalStatus,
    getMonthlyTotals, getSummaryByProjectType, getMultiFYTotals, getConsolidated,
    getAuditLog, getSetting, setSetting, audit, getAllRows,
    getMasterData, getMasterByType, addMasterItem, updateMasterItem, deleteMasterItem, getDynamicConfig,
    exportBudgetCSV, importBudgetCSV, resetDBCache
  };
})();