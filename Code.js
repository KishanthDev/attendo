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
    .setTitle('Modern HRMS Workspace')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// SETUP & INSTALLATION
// ==========================================

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const schemas = {
    [SHEETS.EMP]: ['EmpID', 'Name', 'Email', 'Department', 'Designation', 'Role', 'Status', 'JoiningDate'],
    [SHEETS.ATT]: ['AttID', 'EmpID', 'Date', 'CheckIn', 'CheckOut', 'Hours', 'Status', 'LateArrival'],
    [SHEETS.LEAVE]: ['LeaveID', 'EmpID', 'Type', 'StartDate', 'EndDate', 'Days', 'Status', 'Remarks'],
    [SHEETS.HOL]: ['HolID', 'Name', 'Date', 'Type', 'Description'],
    [SHEETS.SET]: ['Key', 'Value'],
    [SHEETS.AUDIT]: ['LogID', 'Timestamp', 'User', 'Action', 'Details'],
    [SHEETS.ANN]: ['AnnID', 'Date', 'Title', 'Content', 'Status']
  };

  for (const [sheetName, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f8fafc");
    }
  }

  const empSheet = ss.getSheetByName(SHEETS.EMP);
  if (empSheet.getLastRow() === 1) {
    empSheet.appendRow(['EMP-001', 'System Admin', 'admin@company.com', 'IT', 'Administrator', 'ADMIN', 'Active', new Date().toISOString().split('T')[0]]);
    empSheet.appendRow(['EMP-002', 'Demo Employee', 'employee@company.com', 'Engineering', 'Developer', 'EMPLOYEE', 'Active', new Date().toISOString().split('T')[0]]);
  }
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
  
  const rowData = headers.map(h => dataObj[h] !== undefined ? dataObj[h] : '');
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return true;
}

function deleteRowFromSheet(sheetName, keyField, keyValue) {
  const sheet = getDb().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][headers.indexOf(keyField)] === keyValue) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function logAudit(userEmail, action, details) {
  const sheet = getDb().getSheetByName(SHEETS.AUDIT);
  sheet.appendRow([Utilities.getUuid(), new Date().toISOString(), userEmail, action, details]);
}

// ==========================================
// AUTHENTICATION
// ==========================================

function loginUser(inputEmail) {
  let email = inputEmail ? inputEmail.trim() : Session.getActiveUser().getEmail();
  if (!email) email = 'admin@company.com'; 
  
  const employees = getSheetDataAsJSON(SHEETS.EMP);
  let user = employees.find(e => e.Email.toLowerCase() === email.toLowerCase());
  
  if (!user) user = { EmpID: 'UNKNOWN', Name: 'Guest User', Email: email, Role: 'EMPLOYEE', Status: 'Inactive' };
  
  logAudit(email, 'LOGIN', 'User accessed the portal.');
  return user;
}

// ==========================================
// DASHBOARD & ANALYTICS
// ==========================================

function getDashboardData() {
  const emps = getSheetDataAsJSON(SHEETS.EMP);
  const atts = getSheetDataAsJSON(SHEETS.ATT);
  const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
  const hols = getSheetDataAsJSON(SHEETS.HOL);
  const today = new Date().toISOString().split('T')[0];
  
  // Strict requirements: Only count Active EMPLOYEEs, exclude ADMINs
  const activeEmployees = emps.filter(e => e.Status === 'Active' && e.Role === 'EMPLOYEE');
  const activeEmpIds = activeEmployees.map(e => e.EmpID);
  
  const presentToday = atts.filter(a => a.Date.includes(today) && a.CheckIn && activeEmpIds.includes(a.EmpID)).length;
  const leavesToday = leaves.filter(l => l.Status === 'Approved' && l.StartDate <= today && l.EndDate >= today && activeEmpIds.includes(l.EmpID)).length;
  const absentToday = activeEmployees.length - presentToday - leavesToday;
  
  return {
    totalEmployees: activeEmployees.length,
    presentToday: presentToday,
    absentToday: absentToday < 0 ? 0 : absentToday,
    onLeaveToday: leavesToday,
    pendingLeaves: leaves.filter(l => l.Status === 'Pending').length,
    announcements: getSheetDataAsJSON(SHEETS.ANN).filter(a => a.Status === 'Active').slice(-5).reverse(),
    upcomingHolidays: hols.filter(h => h.Date >= today).sort((a,b) => a.Date > b.Date ? 1 : -1).slice(0, 3)
  };
}

// ==========================================
// EMPLOYEE APIs
// ==========================================

function getEmployees() {
  // Requirement: Employee table should never show ADMIN accounts.
  return getSheetDataAsJSON(SHEETS.EMP).filter(e => e.Role === 'EMPLOYEE');
}

function saveEmployee(empData) {
  const isNew = !empData.EmpID;
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';

  if (isNew) {
    empData.EmpID = 'EMP-' + Math.floor(100000 + Math.random() * 900000);
    empData.JoiningDate = new Date().toISOString().split('T')[0];
  }
  
  // Strict Requirement: Backend must enforce Role = EMPLOYEE even if frontend is manipulated
  empData.Role = 'EMPLOYEE';
  if (!empData.Status) empData.Status = 'Active';

  saveRowToSheet(SHEETS.EMP, empData, 'EmpID');
  logAudit(currentUser, isNew ? 'CREATE_EMP' : 'UPDATE_EMP', `Employee ID: ${empData.EmpID}`);
  return { status: 'Success', message: isNew ? 'Employee successfully added.' : 'Employee details updated.' };
}

function deleteEmployee(empID) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  deleteRowFromSheet(SHEETS.EMP, 'EmpID', empID);
  logAudit(currentUser, 'DELETE_EMP', `Deleted Employee ID: ${empID}`);
  return { status: 'Success', message: 'Employee successfully removed.' };
}

// ==========================================
// ATTENDANCE APIs
// ==========================================

function getAttendance() { return getSheetDataAsJSON(SHEETS.ATT); }

function markAttendance(action, empId) {
  const today = new Date().toISOString().split('T')[0];
  const now = new Date();
  const timeStr = now.toTimeString().split(' ')[0];
  const currentUser = Session.getActiveUser().getEmail() || empId;
  
  const sheet = getDb().getSheetByName(SHEETS.ATT);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let rowIndex = -1;
  let checkInTime = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf('EmpID')] === empId && data[i][headers.indexOf('Date')].includes(today)) {
      rowIndex = i + 1; 
      checkInTime = data[i][headers.indexOf('CheckIn')];
      break;
    }
  }
  
  if (action === 'CHECK_IN' && rowIndex === -1) {
    // Check if late (after 09:15 AM)
    const [hours, minutes] = timeStr.split(':');
    const isLate = (parseInt(hours) > 9 || (parseInt(hours) === 9 && parseInt(minutes) > 15)) ? 'Yes' : 'No';
    
    sheet.appendRow(['ATT-' + Date.now(), empId, today, timeStr, '', '', 'Present', isLate]);
    logAudit(currentUser, 'CHECK_IN', `Checked in at ${timeStr}`);
    return { status: 'Success', message: 'Successfully checked in!' };
  } else if (action === 'CHECK_OUT' && rowIndex > -1) {
    // Calculate Hours
    const cIn = new Date(`${today}T${checkInTime}`);
    const workedHours = (Math.abs(now - cIn) / 36e5).toFixed(2);
    
    sheet.getRange(rowIndex, headers.indexOf('CheckOut') + 1).setValue(timeStr);
    sheet.getRange(rowIndex, headers.indexOf('Hours') + 1).setValue(workedHours);
    logAudit(currentUser, 'CHECK_OUT', `Checked out at ${timeStr}. Hours: ${workedHours}`);
    return { status: 'Success', message: 'Successfully checked out!' };
  }
  return { status: 'Error', message: 'Action not valid (Already checked in/out).' };
}

// ==========================================
// LEAVE APIs
// ==========================================

function getLeaves() { return getSheetDataAsJSON(SHEETS.LEAVE); }

function applyLeave(leaveData) {
  const currentUser = Session.getActiveUser().getEmail() || leaveData.EmpID;
  leaveData.LeaveID = 'LV-' + Math.floor(100000 + Math.random() * 900000);
  leaveData.Status = 'Pending';
  
  const start = new Date(leaveData.StartDate);
  const end = new Date(leaveData.EndDate);
  leaveData.Days = Math.ceil(Math.abs(end - start) / (1000 * 60 * 60 * 24)) + 1;
  
  saveRowToSheet(SHEETS.LEAVE, leaveData, 'LeaveID');
  logAudit(currentUser, 'APPLY_LEAVE', `Applied for ${leaveData.Days} days of ${leaveData.Type}`);
  return { status: 'Success', message: 'Leave request successfully submitted!' };
}

function updateLeaveStatus(leaveId, status, remarks = '') {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
  let leave = leaves.find(l => l.LeaveID === leaveId);
  
  if (leave) { 
    leave.Status = status; 
    leave.Remarks = remarks;
    saveRowToSheet(SHEETS.LEAVE, leave, 'LeaveID'); 
    logAudit(currentUser, `LEAVE_${status.toUpperCase()}`, `Leave ${leaveId} ${status}. Remarks: ${remarks}`);
    return { status: 'Success', message: `Leave request has been ${status}.` };
  }
  return { status: 'Error', message: 'Leave request not found.' };
}

// ==========================================
// HOLIDAY APIs
// ==========================================

function getHolidays() { return getSheetDataAsJSON(SHEETS.HOL); }

function saveHoliday(holData) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  if (!holData.HolID) holData.HolID = 'HOL-' + Math.floor(100000 + Math.random() * 900000);
  saveRowToSheet(SHEETS.HOL, holData, 'HolID');
  logAudit(currentUser, 'SAVE_HOLIDAY', `Holiday: ${holData.Name}`);
  return { status: 'Success', message: 'Holiday details saved successfully!' };
}

function deleteHoliday(holId) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  deleteRowFromSheet(SHEETS.HOL, 'HolID', holId);
  logAudit(currentUser, 'DELETE_HOLIDAY', `Removed Holiday ID: ${holId}`);
  return { status: 'Success', message: 'Holiday successfully removed.' };
}

// ==========================================
// ANNOUNCEMENT APIs
// ==========================================

function getAnnouncements() { return getSheetDataAsJSON(SHEETS.ANN); }

function saveAnnouncement(annData) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  if (!annData.AnnID) {
    annData.AnnID = 'ANN-' + Date.now();
    annData.Date = new Date().toISOString().split('T')[0];
    annData.Status = 'Active';
  }
  saveRowToSheet(SHEETS.ANN, annData, 'AnnID');
  logAudit(currentUser, 'SAVE_ANNOUNCEMENT', `Announcement: ${annData.Title}`);
  return { status: 'Success', message: 'Announcement published successfully!' };
}

function deleteAnnouncement(annId) {
  deleteRowFromSheet(SHEETS.ANN, 'AnnID', annId);
  return { status: 'Success', message: 'Announcement deleted.' };
}

function getAuditLogs() { return getSheetDataAsJSON(SHEETS.AUDIT).slice(-100).reverse(); }