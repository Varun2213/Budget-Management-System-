// ============================================================
//  Config.gs — All constants. Edit here to customize.
// ============================================================
const CFG = {
  APP_NAME    : 'Budget Management System',
  DB_NAME     : '💰 BudgetDB (Do Not Edit)',   // hidden sheet file name

  COMPANIES   : ['ABC Realty Pvt Ltd', 'XYZ Developers Pvt Ltd'],
  DEPARTMENTS : ['HR', 'Sales', 'Other'],
  MODULES     : ['Expense', 'Income'],
  PROJECT_TYPES: ['Residential', 'Commercial', 'Plotting'],
  MAIN_HEADS  : ['Other Cost', 'Project Cost', 'Overhead Cost'],

  SUB_HEADS: {
    'Expense': ['Direct Expense', 'Indirect Expense'],
    'Income' : ['Direct Income',  'Indirect Income']
  },

  SUB_SUB_HEADS: {
    'Direct Expense'  : ['Labour Cost','Material Cost','Equipment Cost','Contractor Cost'],
    'Indirect Expense': ['Administration','Marketing & Sales','IT & Technology','Finance Cost'],
    'Direct Income'   : ['Project Revenue','Sales Revenue'],
    'Indirect Income' : ['Interest Income','Other Income']
  },

  CATEGORIES: {
    'Labour Cost'      : ['Wages','Overtime','Contract Labour','ESI/PF','Bonus & Incentives'],
    'Material Cost'    : ['Raw Material','Consumables','Packing Material','Construction Material'],
    'Equipment Cost'   : ['Machinery Hire','Tools & Equipment','Repair & Maintenance'],
    'Contractor Cost'  : ['Civil Contractor','Electrical Contractor','Plumbing','Interior Work'],
    'Administration'   : ['Office Rent','Electricity','Water Charges','Stationery','Telephone & Internet'],
    'Marketing & Sales': ['Advertising','Digital Marketing','Events','Sales Commission','Brochure'],
    'IT & Technology'  : ['Software License','Hardware','AMC Charges','Cloud Services'],
    'Finance Cost'     : ['Bank Charges','Interest on Loan','Professional Fees','Audit Fees','Legal'],
    'Project Revenue'  : ['Residential Projects','Commercial Projects','Plotting Projects'],
    'Sales Revenue'    : ['Unit Sales','Advance Received','Final Payment'],
    'Interest Income'  : ['Bank Interest','FD Returns','Investment Income'],
    'Other Income'     : ['Forfeiture','Penalty Recovery','Misc Income','Rental Income']
  },

  MONTHS      : ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar'],
  ROLES       : ['Finance Admin','HOD','Viewer'],

  STATUS: {
    DRAFT    : 'Draft',
    SUBMITTED: 'Submitted',
    APPROVED : 'Approved',
    LOCKED   : 'Locked',
    REVISION : 'Revision Required'
  },

  // Database sheet tab names (internal — users never see this)
  SHEETS: {
    BUDGET    : 'BudgetEntries',
    ACTUAL    : 'ActualEntries',
    APPROVAL  : 'ApprovalStatus',
    USERS     : 'Users',
    AUDIT     : 'AuditLog',
    SETTINGS  : 'Settings',
    MASTER    : 'MasterData'
  }
};

// Budget entry columns (0-indexed)
const BCOL = {
  ID:0, COMPANY:1, FY:2, DEPT:3, MODULE:4,
  PROJECT_TYPE:5, MAIN_HEAD:6, SUB_HEAD:7, SUB_SUB_HEAD:8, CATEGORY:9,
  SITE_NAME:10, RECURRING:11, NOTES:12,
  APR:13,MAY:14,JUN:15,JUL:16,AUG:17,SEP:18,OCT:19,NOV:20,DEC:21,JAN:22,FEB:23,MAR:24,
  TOTAL:25, STATUS:26, CREATED_BY:27, CREATED_ON:28, MODIFIED_BY:29, MODIFIED_ON:30
};

// Actual entry columns (same structure as Budget)
const ACOL = { ...BCOL };