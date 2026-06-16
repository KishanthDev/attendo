// ==========================================
// CORE SYSTEM CONFIGURATION
// ==========================================

const SHEETS = {
  EMP: 'Employees',
  ATT: 'Attendance',
  LEAVE: 'LeaveRequests',
  HOL: 'Holidays',
  SET: 'Settings',
  AUDIT: 'AuditLogs',
  ANN: 'Announcements'
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HRMS SaaS Workspace')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// SETUP & INSTALLATION
// ==========================================

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const schemas = {
    [SHEETS.EMP]: ['EmpID', 'Name', 'Email', 'Phone', 'Department', 'Designation', 'Manager', 'JoiningDate', 'Role', 'Status'],
    [SHEETS.ATT]: ['AttID', 'EmpID', 'Date', 'CheckIn', 'CheckOut', 'Hours', 'Status'],
    [SHEETS.LEAVE]: ['LeaveID', 'EmpID', 'Type', 'StartDate', 'EndDate', 'Status', 'Remarks'],
    [SHEETS.HOL]: ['HolID', 'Name', 'Date'],
    [SHEETS.SET]: ['Key', 'Value'],
    [SHEETS.AUDIT]: ['LogID', 'Timestamp', 'User', 'Action', 'Details'],
    [SHEETS.ANN]: ['AnnID', 'Date', 'Title', 'Content', 'Author']
  };

  for (const [sheetName, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
    }
  }

  // Insert Sample Users if Employees sheet is empty
  const empSheet = ss.getSheetByName(SHEETS.EMP);
  if (empSheet.getLastRow() === 1) {
    empSheet.appendRow(['EMP-001', 'System Admin', 'admin@company.com', '1234567890', 'IT', 'Admin', 'None', new Date().toISOString().split('T')[0], 'ADMIN', 'Active']);
    empSheet.appendRow(['EMP-002', 'Demo Employee', 'employee@company.com', '0987654321', 'Engineering', 'Developer', 'EMP-001', new Date().toISOString().split('T')[0], 'EMPLOYEE', 'Active']);
  }
  
  logAudit('System', 'INITIALIZE', 'HRMS Setup Completed');
  return "Setup successful.";
}

// ==========================================
// UTILITIES & DB HELPERS
// ==========================================

function getDb() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheetDataAsJSON(sheetName) {
  const sheet = getDb().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data.shift();
  return data.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] instanceof Date ? row[i].toISOString() : row[i];
    });
    return obj;
  });
}

function saveRowToSheet(sheetName, dataObj, keyField) {
  const sheet = getDb().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyValue = dataObj[keyField];
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf(keyField)] === keyValue) {
      rowIndex = i + 1;
      break;
    }
  }
  
  const rowData = headers.map(h => dataObj[h] || '');
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return true;
}

function logAudit(userEmail, action, details) {
  const sheet = getDb().getSheetByName(SHEETS.AUDIT);
  sheet.appendRow([Utilities.getUuid(), new Date().toISOString(), userEmail, action, details]);
}

// ==========================================
// AUTHENTICATION & SESSION
// ==========================================

function loginUser(inputEmail) {
  // Use input email for testing role switching, fallback to active Workspace user
  let email = inputEmail ? inputEmail.trim() : Session.getActiveUser().getEmail();
  if (!email) email = 'admin@company.com'; // Fallback if tested outside domain
  
  const employees = getSheetDataAsJSON(SHEETS.EMP);
  let user = employees.find(e => e.Email.toLowerCase() === email.toLowerCase());
  
  if (!user) {
    user = { EmpID: 'UNKNOWN', Name: 'Guest User', Email: email, Role: 'EMPLOYEE', Status: 'Inactive' };
  }
  
  logAudit(email, 'LOGIN', 'User accessed application via login gateway');
  return user;
}

// ==========================================
// ADMIN & DASHBOARD APIs
// ==========================================

function getDashboardData() {
  const emps = getSheetDataAsJSON(SHEETS.EMP);
  const atts = getSheetDataAsJSON(SHEETS.ATT);
  const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
  
  const today = new Date().toISOString().split('T')[0];
  
  const presentToday = atts.filter(a => a.Date.includes(today) && a.CheckIn).length;
  const leavesToday = leaves.filter(l => l.Status === 'Approved' && l.StartDate <= today && l.EndDate >= today).length;
  const pendingLeaves = leaves.filter(l => l.Status === 'Pending').length;
  
  return {
    totalEmployees: emps.filter(e => e.Status === 'Active').length,
    presentToday: presentToday,
    onLeaveToday: leavesToday,
    pendingLeaves: pendingLeaves,
    recentLeaves: leaves.slice(-5).reverse(),
    announcements: getSheetDataAsJSON(SHEETS.ANN).slice(-3).reverse()
  };
}

function getEmployees() { return getSheetDataAsJSON(SHEETS.EMP); }
function saveEmployee(empData) {
  if (!empData.EmpID) empData.EmpID = 'EMP-' + Date.now().toString().slice(-6);
  saveRowToSheet(SHEETS.EMP, empData, 'EmpID');
  logAudit(Session.getActiveUser().getEmail(), 'SAVE_EMP', `Employee updated: ${empData.EmpID}`);
  return getEmployees();
}
function deleteEmployee(empID) { return getEmployees(); }

// ==========================================
// ATTENDANCE APIs
// ==========================================

function getAttendance() { return getSheetDataAsJSON(SHEETS.ATT); }

function markAttendance(action, empId) {
  const today = new Date().toISOString().split('T')[0];
  const now = new Date();
  const timeStr = now.toTimeString().split(' ')[0];
  
  const sheet = getDb().getSheetByName(SHEETS.ATT);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let rowIndex = -1;
  let currentRecord = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('EmpID')] === empId && data[i][headers.indexOf('Date')].includes(today)) {
      rowIndex = i + 1;
      currentRecord = data[i];
      break;
    }
  }
  
  if (action === 'CHECK_IN' && rowIndex === -1) {
    const attId = 'ATT-' + Date.now();
    sheet.appendRow([attId, empId, today, timeStr, '', '', 'Present']);
    return { status: 'Success', message: 'Checked In Successfully' };
  } else if (action === 'CHECK_OUT' && rowIndex > -1) {
    const checkInTime = new Date(`${today}T${currentRecord[headers.indexOf('CheckIn')]}`);
    const hours = Math.abs(now - checkInTime) / 36e5;
    
    sheet.getRange(rowIndex, headers.indexOf('CheckOut') + 1).setValue(timeStr);
    sheet.getRange(rowIndex, headers.indexOf('Hours') + 1).setValue(hours.toFixed(2));
    return { status: 'Success', message: 'Checked Out Successfully' };
  }
  
  return { status: 'Error', message: 'Invalid action or already checked in/out' };
}

// ==========================================
// LEAVE & MISC APIs
// ==========================================

function getLeaves() { return getSheetDataAsJSON(SHEETS.LEAVE); }
function applyLeave(leaveData) {
  leaveData.LeaveID = 'LV-' + Date.now().toString().slice(-6);
  leaveData.Status = 'Pending';
  saveRowToSheet(SHEETS.LEAVE, leaveData, 'LeaveID');
  return getLeaves();
}
function updateLeaveStatus(leaveId, status, remarks) {
  const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
  let leave = leaves.find(l => l.LeaveID === leaveId);
  if (leave) {
    leave.Status = status;
    leave.Remarks = remarks;
    saveRowToSheet(SHEETS.LEAVE, leave, 'LeaveID');
  }
  return getLeaves();
}

function getHolidays() { return getSheetDataAsJSON(SHEETS.HOL); }
function saveHoliday(holData) {
  if (!holData.HolID) holData.HolID = 'HOL-' + Date.now();
  saveRowToSheet(SHEETS.HOL, holData, 'HolID');
  return getHolidays();
}
function getAnnouncements() { return getSheetDataAsJSON(SHEETS.ANN); }