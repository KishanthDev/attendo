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

function debugTodaySummary() {
  Logger.log(
    JSON.stringify(
      getAttendanceSummary("EMP-204415"),
      null,
      2
    )
  );
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Modern HRMS Workspace')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function debugRawCheckIn() {
  const sheet = getDb().getSheetByName(SHEETS.ATT);
  const data = sheet.getDataRange().getValues();

  Logger.log(typeof data[1][3]);
  Logger.log(data[1][3]);
}
// ==========================================
// SETUP & INSTALLATION
// ==========================================

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemas = {
    [SHEETS.EMP]: ['EmpID', 'Name', 'Email', 'Password', 'Department', 'Designation', 'Role', 'Status', 'JoiningDate'],
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

  // Default Settings
  const setSheet = ss.getSheetByName(SHEETS.SET);
  if (setSheet.getLastRow() === 1) {
    setSheet.appendRow(['DailyWorkingHours', '8']);
    setSheet.appendRow(['LateArrivalTime', '09:15']);
  }
  // Sample Users
  const empSheet = ss.getSheetByName(SHEETS.EMP);

  const data = empSheet.getDataRange().getValues();

  const adminExists = data.some(
    row => row[2] === 'admin@company.com'
  );

  if (!adminExists) {
    empSheet.appendRow([
      'EMP001',
      'System Admin',
      'admin@company.com',
      'admin123',
      'Administration',
      'Administrator',
      'ADMIN',
      'Active',
      new Date()
    ]);

    empSheet.appendRow([
      'EMP002',
      'John Doe',
      'john@company.com',
      'john123',
      'Engineering',
      'Software Engineer',
      'EMPLOYEE',
      'Active',
      new Date()
    ]);

    empSheet.appendRow([
      'EMP003',
      'Jane Smith',
      'jane@company.com',
      'jane123',
      'HR',
      'HR Executive',
      'EMPLOYEE',
      'Active',
      new Date()
    ]);
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
    const obj = {};

    headers.forEach((header, i) => {
      const value = row[i];

      // Attendance Date column
      if (sheetName === SHEETS.ATT && header === 'Date') {
        if (value instanceof Date) {
          obj[header] = Utilities.formatDate(
            value,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd"
          );
        } else {
          obj[header] = String(value);
        }
        return;
      }

      // Check In / Check Out
      if (header === 'CheckIn' || header === 'CheckOut') {
        if (value instanceof Date) {
          obj[header] = String(value.getTime());
        } else {
          obj[header] = value || '';
        }
        return;
      }

      // Settings -> Late Arrival Time
      if (
        sheetName === SHEETS.SET &&
        row[0] === 'LateArrivalTime' &&
        header === 'Value'
      ) {
        if (value instanceof Date) {
          obj[header] = Utilities.formatDate(
            value,
            Session.getScriptTimeZone(),
            "HH:mm"
          );
        } else {
          obj[header] = String(value).trim();
        }
        return;
      }

      // Generic Date Handling
      if (value instanceof Date) {
        obj[header] = Utilities.formatDate(
          value,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
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

/**
 * Reusable helper to check if a given date is a holiday or a weekend (Sunday).
 * @param {Date} date The date to check.
 * @param {Array<Object>} holidays An array of holiday objects from the sheet.
 * @returns {Object|null} The holiday object if it's a holiday/weekend, otherwise null.
 */
function isDateHoliday(date, holidays) {
  if (!date || !(date instanceof Date)) return null;
  if (date.getDay() === 0) { // Sunday is a weekend
    return { Name: 'Weekend', Type: 'Weekend', Description: 'Sunday Weekend' };
  }
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return holidays.find(h => h.Date === dateStr) || null;
}

// ==========================================
// SETTINGS
// ==========================================

function getAppSettings() {
  const data = getSheetDataAsJSON(SHEETS.SET);

  let settings = {
    DailyWorkingHours: 8,
    LateArrivalTime: "09:15"
  };

  data.forEach(row => {
    const key = String(row.Key).trim();

    if (key === 'DailyWorkingHours') {
      settings.DailyWorkingHours = Number(row.Value);
    }

    if (row.Key === 'LateArrivalTime') {
      Logger.log(row.Value);

      if (row.Value instanceof Date) {
        settings.LateArrivalTime = Utilities.formatDate(
          row.Value,
          Session.getScriptTimeZone(),
          'HH:mm'
        );
      } else {
        settings.LateArrivalTime = String(row.Value);
      }
    }
  });

  return settings;
}

function saveAppSettings(settingsData) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  saveRowToSheet(SHEETS.SET, { Key: 'DailyWorkingHours', Value: settingsData.DailyWorkingHours }, 'Key');
  saveRowToSheet(SHEETS.SET, { Key: 'LateArrivalTime', Value: settingsData.LateArrivalTime }, 'Key');
  logAudit(currentUser, 'UPDATE_SETTINGS', `Working Hours: ${settingsData.DailyWorkingHours}, Late Time: ${settingsData.LateArrivalTime}`);
  return { status: 'Success', message: 'Settings updated successfully!' };
}

// ==========================================
// AUTHENTICATION
// ==========================================

function loginUser(email, password) {
  if (!email || !password) {
    return { status: 'Error', message: 'Email and password are required.' };
  }

  const employees = getSheetDataAsJSON(SHEETS.EMP);
  let user = employees.find(e => e.Email.toLowerCase() === email.trim().toLowerCase());

  if (!user) {
    return { status: 'Error', message: 'Account not found.' };
  }

  if (user.Password !== password) {
    return { status: 'Error', message: 'Invalid credentials.' };
  }

  if (user.Status !== 'Active') {
    return { status: 'Error', message: 'Your account is inactive.' };
  }

  // Strip password before returning to client for security
  delete user.Password;

  logAudit(email, 'LOGIN', 'User accessed the portal.');
  return { status: 'Success', user: user };
}

// ==========================================
// DASHBOARD
// ==========================================

function getDashboardData() {
  const todayDate = new Date();
  const todayStr = Utilities.formatDate(todayDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const emps = getSheetDataAsJSON(SHEETS.EMP);
  const atts = getSheetDataAsJSON(SHEETS.ATT);
  const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
  const hols = getSheetDataAsJSON(SHEETS.HOL);

  // --- Holiday Calculations ---
  const todayHoliday = isDateHoliday(todayDate, hols);
  const upcomingHolidays = hols
    .filter(h => h.Date >= todayStr)
    .sort((a, b) => a.Date > b.Date ? 1 : -1)
    .map(h => {
      const holDate = new Date(h.Date);
      // Adjust for timezone differences when calculating days
      const todayMidnight = new Date(todayStr);
      const holMidnight = new Date(holDate.getFullYear(), holDate.getMonth(), holDate.getDate());
      const diffTime = holMidnight.getTime() - todayMidnight.getTime();
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      h.daysLeft = diffDays;
      return h;
    });

  const nextHoliday = upcomingHolidays.find(h => h.Type !== 'Weekend') || null;

  // --- Attendance Calculations ---
  const activeEmployees = emps.filter(e => e.Status === 'Active' && e.Role === 'EMPLOYEE');
  const activeEmpIds = activeEmployees.map(e => e.EmpID);

  const presentSet = new Set();
  atts.forEach(a => {
    if (a.CheckIn && activeEmpIds.includes(a.EmpID)) {
      let rStr = String(a.Date).split('T')[0];
      if (rStr === todayStr) presentSet.add(a.EmpID);
    }
  });

  const leavesTodayCount = leaves.filter(l => l.Status === 'Approved' && l.StartDate <= todayStr && l.EndDate >= todayStr && activeEmpIds.includes(l.EmpID)).length;

  // Don't count attendance metrics if it's a holiday
  const presentToday = todayHoliday ? 0 : presentSet.size;
  const absentToday = todayHoliday ? 0 : (activeEmployees.length - presentToday - leavesTodayCount);

  return {
    totalEmployees: activeEmployees.length,
    presentToday: presentToday,
    absentToday: absentToday < 0 ? 0 : absentToday,
    onLeaveToday: leavesTodayCount,
    pendingLeaves: leaves.filter(l => l.Status === 'Pending').length,
    announcements: getSheetDataAsJSON(SHEETS.ANN).filter(a => a.Status === 'Active').slice(-5).reverse(),

    // New & Enhanced Holiday Data
    todayHoliday: todayHoliday,
    nextHoliday: nextHoliday,
    upcomingHolidays: upcomingHolidays.slice(0, 3) // Keep it to 3 for the widget
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
// ATTENDANCE APIs (RE-ENGINEERED)
// ==========================================

function getAttendanceSummary(empIdFilter = null) {
  const rawData = getSheetDataAsJSON(SHEETS.ATT);
  const settings = getAppSettings();
  const reqHours = Number(settings.DailyWorkingHours) || 8;
  const holidays = getSheetDataAsJSON(SHEETS.HOL); // Get all holidays

  let filtered = rawData;
  if (empIdFilter) {
    filtered = rawData.filter(r => r.EmpID === empIdFilter);
  }

  // Group by EmpID + Date
  const groups = {};
  filtered.forEach(row => {
    const dStr = row.Date; // already yyyy-MM-dd

    const key = row.EmpID + "_" + dStr;

    if (!groups[key]) {
      groups[key] = {
        EmpID: row.EmpID,
        Date: dStr,
        TotalHours: 0,
        Sessions: [],
        RequiredHours: reqHours
      };
    }

    groups[key].Sessions.push(row);

    const h = Number(row.Hours);
    if (!isNaN(h)) {
      groups[key].TotalHours += h;
    }
  });

  const result = Object.values(groups).map(g => {

    const hasActiveSession = g.Sessions.some(
      s => !s.CheckOut || s.CheckOut === ""
    );

    if (hasActiveSession) {
      g.Status = 'Present';
    }
    else if (g.TotalHours >= g.RequiredHours) {
      g.Status = 'Completed';
    }
    else if (g.TotalHours >= (g.RequiredHours / 2)) {
      g.Status = 'Half Day';
    }
    else if (g.TotalHours > 0) {
      g.Status = 'Present';
    }
    else {
      g.Status = 'Absent';
    }

    g.Sessions.sort((a, b) =>
      Number(String(a.CheckIn).replace(/,/g, '')) -
      Number(String(b.CheckIn).replace(/,/g, ''))
    );

    return g;
  });

  result.sort((a, b) => a.Date > b.Date ? -1 : 1);
  return result;
}

function markAttendance(action, empId) {
  const now = new Date();
  const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  const epochTimestamp = now.getTime();
  const currentUser = Session.getActiveUser().getEmail() || empId;
  const settings = getAppSettings();

  const sheet = getDb().getSheetByName(SHEETS.ATT);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  let lastRecordIndex = -1;
  let lastCheckIn = null;
  let lastCheckOut = null;
  let lastDateStr = null;
  let sessionsToday = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    let rowEmpId = data[i][headers.indexOf('EmpID')];
    if (rowEmpId === empId) {
      let rawDate = data[i][headers.indexOf('Date')];
      let d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        let rStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        if (rStr === todayStr) sessionsToday++;

        if (lastRecordIndex === -1) {
          lastRecordIndex = i + 1;
          lastCheckIn = data[i][headers.indexOf('CheckIn')];
          lastCheckOut = data[i][headers.indexOf('CheckOut')];
          lastDateStr = rStr;
        }
      }
    }
  }

  const hasActiveSessionToday = (lastRecordIndex > -1 && lastDateStr === todayStr && (!lastCheckOut || lastCheckOut === ""));

  if (action === 'CHECK_IN') {
    if (hasActiveSessionToday) return { status: 'Error', message: 'You have an active session! Check out first.' };

    let isLate = 'No';
    if (sessionsToday === 0) { // Only check late arrival on the FIRST check-in of the day
      const [lateH, lateM] = (settings.LateArrivalTime || "09:15").split(':').map(Number);
      const currentH = now.getHours();
      const currentM = now.getMinutes();
      if (currentH > lateH || (currentH === lateH && currentM > lateM)) isLate = 'Yes';
    }
    Logger.log('sessionsToday = ' + sessionsToday);
    Logger.log('isLate = ' + isLate);
    Logger.log([
      'ATT-' + epochTimestamp,
      empId,
      todayStr,
      String(epochTimestamp),
      "",
      "",
      'Present',
      isLate
    ]);

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
    if (!hasActiveSessionToday) return { status: 'Error', message: 'No active Check-In found. You must Check In first.' };

    const cleanCheckInStr = String(lastCheckIn).replace(/,/g, '');
    const cInEpoch = Number(cleanCheckInStr);
    let workedHours = "0.00";
    if (!isNaN(cInEpoch) && cInEpoch > 0) {
      workedHours = (Math.abs(epochTimestamp - cInEpoch) / 36e5).toFixed(2);
    }

    sheet.getRange(
      lastRecordIndex,
      headers.indexOf('CheckOut') + 1
    ).setValue(String(epochTimestamp));
    sheet.getRange(lastRecordIndex, headers.indexOf('Hours') + 1).setValue(workedHours);
    logAudit(currentUser, 'CHECK_OUT', `Checked out. Session Hours: ${workedHours}`);

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