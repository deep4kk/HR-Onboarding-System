/**
 * HR Onboarding & Employee Database System
 */
const CONFIG = { SHEET_NAMES: { USERS: 'Users', EMPLOYEES: 'Employees', ONBOARDING_CHECKLISTS: 'OnboardingChecklists', DOCUMENTS: 'Documents', AUDIT_LOG: 'AuditLog' }, CHECKLIST_ITEMS: ['Offer Letter', 'Employment Contract', 'ID Proof', 'Address Proof', 'Bank Details', 'PF Enrollment', 'Health Insurance', 'Laptop Setup', 'Email Setup', 'Team Introduction', 'Policy Acknowledgment', 'First Day Orientation'] };
const SHEET_CONFIG = const SHEET_CONFIG = {
  Users: { columns: ['Email', 'Name', 'Role', 'Department', 'Active'] },
  Employees: { columns: ['EmployeeID', 'FirstName', 'LastName', 'Email', 'Phone', 'Department', 'Designation', 'JoinDate', 'Status', 'Manager', 'CreatedAt'] },
  OnboardingChecklists: { columns: ['ChecklistID', 'EmployeeID', 'Item', 'Status', 'CompletedBy', 'CompletedAt', 'Notes'] },
  Documents: { columns: ['DocID', 'EmployeeID', 'DocType', 'DocName', 'UploadDate', 'FileURL'] },
  AuditLog: { columns: ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue'] }
};;

// ============== ENHANCED UTILITIES (v2.0) ==============
const VERSION = '2.0.0';

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  for (const [sheetName, config] of Object.entries(SHEET_CONFIG)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      created.push(sheetName);
      sheet.getRange(1, 1, 1, config.columns.length).setValues([config.columns]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      if (config.sampleData) config.sampleData.forEach(row => sheet.appendRow(row));
    }
  }
  Logger.log('InitializeSheets: Created ' + created.join(', '));
  return { created };
}

function handleError(error, context) {
  const errorMsg = error instanceof Error ? error.message : String(error);
  Logger.log('[ERROR] ' + context + ': ' + errorMsg);
  if (typeof logAction === 'function') logAction('ERROR', context, '', errorMsg);
  return { success: false, error: errorMsg };
}

function backupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = 'Backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  ss.copy(backupName);
  if (typeof logAction === 'function') logAction('BACKUP_CREATED', 'System', '', backupName);
  return { success: true, backupName };
}

function exportToPDF(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  const pdf = DriveApp.getFileById(ss.getId()).getAs('application/pdf');
  const pdfName = sheet.getName() + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf';
  DriveApp.getRootFolder().createFile(pdf).setName(pdfName);
  return { success: true, fileName: pdfName };
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 System Menu')
    .addItem('📊 Initialize Sheets', 'initializeSheets')
    .addItem('💾 Create Backup', 'backupData')
    .addItem('📄 Export to PDF', 'exportToPDF')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('System v' + VERSION, 'Enhanced with:\n- initializeSheets()\n- backupData()\n- exportToPDF()', ui.ButtonSet.OK);
}

function getSheet(name) { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName(name); if (!sheet) { sheet = ss.insertSheet(name); setupHeaders(sheet, name); } return sheet; }
function setupHeaders(sheet, name) { const h = { Users: ['Email', 'Name', 'Role', 'Department', 'Active'], Employees: ['EmployeeID', 'FirstName', 'LastName', 'Email', 'Phone', 'Department', 'Designation', 'JoinDate', 'Status', 'Manager', 'CreatedAt'], OnboardingChecklists: ['ChecklistID', 'EmployeeID', 'Item', 'Status', 'CompletedBy', 'CompletedAt', 'Notes'], Documents: ['DocID', 'EmployeeID', 'DocType', 'DocName', 'UploadDate', 'FileURL'], AuditLog: ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue'] }; if (h[name]) { sheet.getRange(1, 1, 1, h[name].length).setValues([h[name]]); sheet.getRange(1, 1, 1, h[name].length).setFontWeight('bold'); } }
function generateId(prefix) { return prefix + '-' + new Date().getFullYear() + '-' + Utilities.getUuid().substring(0, 6).toUpperCase(); }
function getCurrentUser() { const email = Session.getActiveUser().getEmail(); const sheet = getSheet(CONFIG.SHEET_NAMES.USERS); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0].toLowerCase() === email.toLowerCase()) return { email: data[i][0], name: data[i][1], role: data[i][2] }; } return { email, name: 'Unknown', role: 'Employee' }; }
function requireRole(roles) { const user = getCurrentUser(); if (!roles.includes(user.role)) throw new Error('Unauthorized'); return user; }
function logAction(action, recordId, oldValue, newValue) { const sheet = getSheet(CONFIG.SHEET_NAMES.AUDIT_LOG); sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, recordId, oldValue, newValue]); }
function addEmployee(firstName, lastName, email, phone, department, designation, joinDate, manager) { requireRole(['HR Admin', 'Admin']); const employeeID = generateId('EMP'); const sheet = getSheet(CONFIG.SHEET_NAMES.EMPLOYEES); sheet.appendRow([employeeID, firstName, lastName, email, phone, department, designation, new Date(joinDate), 'Onboarding', manager || '', new Date()]); const checklistSheet = getSheet(CONFIG.SHEET_NAMES.ONBOARDING_CHECKLISTS); CONFIG.CHECKLIST_ITEMS.forEach(item => { checklistSheet.appendRow([generateId('CHK'), employeeID, item, 'Pending', '', '', '']); }); logAction('ADD_EMPLOYEE', employeeID, '', firstName + ' ' + lastName); return { success: true, employeeID }; }
function updateChecklistItem(checklistID, status, notes) { const user = getCurrentUser(); const sheet = getSheet(CONFIG.SHEET_NAMES.ONBOARDING_CHECKLISTS); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === checklistID) { const row = i + 1; sheet.getRange(row, 4, 1, 1).setValue(status); sheet.getRange(row, 5, 1, 1).setValue(user.email); sheet.getRange(row, 6, 1, 1).setValue(new Date()); sheet.getRange(row, 7, 1, 1).setValue(notes || ''); logAction('UPDATE_CHECKLIST', checklistID, '', status); updateEmployeeOnboardingStatus(data[i][1]); return { success: true }; } } throw new Error('Item not found'); }
function updateEmployeeOnboardingStatus(employeeID) { const sheet = getSheet(CONFIG.SHEET_NAMES.ONBOARDING_CHECKLISTS); const data = sheet.getDataRange().getValues(); let total = 0, completed = 0; for (let i = 1; i < data.length; i++) { if (data[i][1] === employeeID) { total++; if (data[i][3] === 'Completed') completed++; } } const empSheet = getSheet(CONFIG.SHEET_NAMES.EMPLOYEES); const empData = empSheet.getDataRange().getValues(); for (let i = 1; i < empData.length; i++) { if (empData[i][0] === employeeID) { const row = i + 1; const status = completed === total ? 'Active' : 'Onboarding'; empSheet.getRange(row, 9, 1, 1).setValue(status); return; } } }
function getAllEmployees() { const sheet = getSheet(CONFIG.SHEET_NAMES.EMPLOYEES); const data = sheet.getDataRange().getValues(); const employees = []; for (let i = 1; i < data.length; i++) { employees.push({ employeeID: data[i][0], firstName: data[i][1], lastName: data[i][2], email: data[i][3], phone: data[i][4], department: data[i][5], designation: data[i][6], joinDate: data[i][7], status: data[i][8], manager: data[i][9] }); } return employees; }
function getEmployeeChecklist(employeeID) { const sheet = getSheet(CONFIG.SHEET_NAMES.ONBOARDING_CHECKLISTS); const data = sheet.getDataRange().getValues(); const checklist = []; for (let i = 1; i < data.length; i++) { if (data[i][1] === employeeID) checklist.push({ checklistID: data[i][0], item: data[i][2], status: data[i][3], completedBy: data[i][4], completedAt: data[i][5], notes: data[i][6] }); } return checklist; }
function getDashboardStats() { const employees = getAllEmployees(); const active = employees.filter(e => e.status === 'Active').length; const onboarding = employees.filter(e => e.status === 'Onboarding').length; const departments = {}; employees.forEach(e => { departments[e.department] = (departments[e.department] || 0) + 1; }); return { totalEmployees: employees.length, active, onboarding, departments, employees }; }
function doGet(e) { return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('HR Onboarding').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
