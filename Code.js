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

 Object.entries(schemas).forEach(([sheetName, headers]) => {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  });
  return "Setup successful.";
}

function testCheckInCell() {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("Attendance");

  Logger.log(
    sheet.getRange(2,4).getValue()
  );
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
    const value = row[i];

    // Handle CheckIn / CheckOut separately
    if (header === 'CheckIn' || header === 'CheckOut') {

      if (value instanceof Date) {
        // If Sheets converted epoch to Date, convert back safely
        obj[header] = isNaN(value.getTime())
          ? ''
          : String(value.getTime());
      } else {
        obj[header] = value || '';
      }

      return;
    }

    // Handle all date fields safely
    if (value instanceof Date) {
      obj[header] = isNaN(value.getTime())
        ? ''
        : value.toISOString();
    } else {
      obj[header] = value;
    }
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
// DASHBOARD APIs
// ==========================================

function getDashboardData() {
  const emps = getSheetDataAsJSON(SHEETS.EMP);
  const atts = getSheetDataAsJSON(SHEETS.ATT);
  const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
  const hols = getSheetDataAsJSON(SHEETS.HOL);
  
  const todayDate = new Date();
  const todayStr = `${todayDate.getFullYear()}-${String(todayDate.getMonth()+1).padStart(2,'0')}-${String(todayDate.getDate()).padStart(2,'0')}`;
  
  const activeEmployees = emps.filter(e => e.Status === 'Active' && e.Role === 'EMPLOYEE');
  const activeEmpIds = activeEmployees.map(e => e.EmpID);
  
  const presentToday = atts.filter(a => {
    if (!a.CheckIn || !activeEmpIds.includes(a.EmpID)) return false;
    let d = new Date(a.Date);
    if (isNaN(d.getTime())) return false;
    let rStr = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    return rStr === todayStr;
  }).length;
  
  const leavesToday = leaves.filter(l => l.Status === 'Approved' && l.StartDate <= todayStr && l.EndDate >= todayStr && activeEmpIds.includes(l.EmpID)).length;
  const absentToday = activeEmployees.length - presentToday - leavesToday;
  
  return {
    totalEmployees: activeEmployees.length,
    presentToday: presentToday,
    absentToday: absentToday < 0 ? 0 : absentToday,
    onLeaveToday: leavesToday,
    pendingLeaves: leaves.filter(l => l.Status === 'Pending').length,
    announcements: getSheetDataAsJSON(SHEETS.ANN).filter(a => a.Status === 'Active').slice(-5).reverse(),
    upcomingHolidays: hols.filter(h => h.Date >= todayStr).sort((a,b) => a.Date > b.Date ? 1 : -1).slice(0, 3)
  };
}

// ==========================================
// EMPLOYEE APIs
// ==========================================

function getEmployees() {
  return getSheetDataAsJSON(SHEETS.EMP).filter(e => e.Role === 'EMPLOYEE');
}

function saveEmployee(empData) {
  const isNew = !empData.EmpID;
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';

  if (isNew) {
    empData.EmpID = 'EMP-' + Math.floor(100000 + Math.random() * 900000);
    empData.JoiningDate = new Date().toISOString().split('T')[0];
  }
  
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
// ATTENDANCE APIs (FIXED: DAILY SESSION LOCK)
// ==========================================

function getAttendance() { return getSheetDataAsJSON(SHEETS.ATT); }

function markAttendance(action, empId) {
  const now = new Date();
  const todayStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
  const epochTimestamp = now.getTime();
  const currentUser = Session.getActiveUser().getEmail() || empId;
  
  const sheet = getDb().getSheetByName(SHEETS.ATT);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let lastRecordIndex = -1;
  let lastCheckIn = null;
  let lastCheckOut = null;
  let lastDateStr = null;

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][headers.indexOf('EmpID')] === empId) {
      lastRecordIndex = i + 1; 
      lastCheckIn = data[i][headers.indexOf('CheckIn')];
      lastCheckOut = data[i][headers.indexOf('CheckOut')];
      
      let rawDate = data[i][headers.indexOf('Date')];
      let d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        lastDateStr = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      }
      break;
    }
  }
  
  // Strict Scope: Only block if they have an active session TODAY.
  const hasActiveSessionToday = (lastRecordIndex > -1 && lastDateStr === todayStr && (!lastCheckOut || lastCheckOut === ""));

  if (action === 'CHECK_IN') {
    if (hasActiveSessionToday) {
      return { status: 'Error', message: 'You already have an active check-in today! You must check out first.' };
    }
    
    const hours = now.getHours();
    const minutes = now.getMinutes();
    const isLate = (hours > 9 || (hours === 9 && minutes > 15)) ? 'Yes' : 'No';
    
    sheet.appendRow([
  'ATT-' + epochTimestamp,
  empId,
  todayStr,
  String(epochTimestamp),
  "",
  "",
  'Present',
  isLate
]);
    logAudit(currentUser, 'CHECK_IN', `Checked in at Epoch ${epochTimestamp}`);
    return { status: 'Success', message: 'Successfully checked in!' };
    
  } else if (action === 'CHECK_OUT') {
    if (!hasActiveSessionToday) {
      return { status: 'Error', message: 'No active Check-In found for today. You must Check In first.' };
    }
    
    // Safely parse check in string, stripping commas
    const cleanCheckInStr = String(lastCheckIn).replace(/,/g, '');
    const cInEpoch = Number(cleanCheckInStr);
    
    let workedHours = "0.00";
    if (!isNaN(cInEpoch) && cInEpoch > 0) {
      workedHours = (Math.abs(epochTimestamp - cInEpoch) / 36e5).toFixed(2);
    }
    
    sheet.getRange(lastRecordIndex, headers.indexOf('CheckOut') + 1).setValue(epochTimestamp);
    sheet.getRange(lastRecordIndex, headers.indexOf('Hours') + 1).setValue(workedHours);
    logAudit(currentUser, 'CHECK_OUT', `Checked out. Hours: ${workedHours}`);
    
    return { status: 'Success', message: 'Successfully checked out!' };
  }
  
  return { status: 'Error', message: 'Invalid action payload.' };
}

// ==========================================
// LEAVE APIs
// ==========================================

function getLeaves() { return getSheetDataAsJSON(SHEETS.LEAVE); }

function applyLeave(leaveData) {
  const currentUser = Session.getActiveUser().getEmail() || leaveData.EmpID;
  leaveData.LeaveID = 'LV-' + Math.floor(100000 + Math.random() * 900000);
  leaveData.Status = 'Pending';
  leaveData.StartDate = leaveData.Date;
  leaveData.EndDate = leaveData.Date;
  leaveData.Days = 1;
  saveRowToSheet(SHEETS.LEAVE, leaveData, 'LeaveID');
  logAudit(currentUser, 'APPLY_LEAVE', `Applied for ${leaveData.Days} day of ${leaveData.Type}`);
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