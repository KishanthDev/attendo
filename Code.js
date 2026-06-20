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
  ANN: 'Announcements',
  DOCUMENTS: 'Documents',
  CLIENTS: 'Clients',
  CLIENT_ASSIGNMENTS: 'ClientAssignments',
  SKILLS: 'Skills',
  EMP_SKILLS: 'EmployeeSkills'
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
    [SHEETS.EMP]: ['EmpID', 'Name', 'Email', 'Password', 'Department', 'Designation', 'Role', 'Status', 'JoiningDate', 'ExitDate', 'MustChangePassword'],
    [SHEETS.ATT]: ['AttID', 'EmpID', 'Date', 'CheckIn', 'CheckOut', 'Hours', 'Status', 'LateArrival'],
    [SHEETS.LEAVE]: ['LeaveID', 'EmpID', 'Type', 'StartDate', 'EndDate', 'Days', 'Status', 'Remarks'],
    [SHEETS.HOL]: ['HolID', 'Name', 'Date', 'Type', 'Description'],
    [SHEETS.SET]: ['Key', 'Value'],
    [SHEETS.AUDIT]: ['LogID', 'Timestamp', 'User', 'Action', 'Details'],
    [SHEETS.ANN]: ['AnnID', 'Date', 'Title', 'Content', 'Status'],
    [SHEETS.DOCUMENTS]: ['DocID', 'EmpID', 'FileName', 'DocumentType', 'Month', 'Year', 'DriveFileID', 'UploadDate', 'UploadedBy'],
    [SHEETS.CLIENTS]: ['ClientID', 'ClientName', 'WorkingHours', 'Technologies', 'Status', 'StartDate', 'EndDate', 'Description', 'CreatedAt'],
    [SHEETS.CLIENT_ASSIGNMENTS]: ['AssignmentID', 'ClientID', 'EmpID', 'EmployeeName', 'AssignedDate', 'Status'],
    [SHEETS.SKILLS]: ['SkillID', 'Category', 'SkillName', 'Status'],
    [SHEETS.EMP_SKILLS]: ['EmpID', 'SkillID', 'SkillName', 'Category', 'UpdatedAt']
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
    setSheet.appendRow(['CasualLeaveQuota', '10']);
    setSheet.appendRow(['OptionalLeaveQuota', '2']);
  }

  // Seed Default Skills if empty
  const skillsSheet = ss.getSheetByName(SHEETS.SKILLS);
  if (skillsSheet.getLastRow() === 1) {
    const defaultSkills = [
      ['SKL-101', 'Languages', 'JavaScript', 'Active'],
      ['SKL-102', 'Languages', 'TypeScript', 'Active'],
      ['SKL-103', 'Languages', 'Python', 'Active'],
      ['SKL-201', 'Frameworks', 'React', 'Active'],
      ['SKL-202', 'Frameworks', 'Next.js', 'Active'],
      ['SKL-203', 'Frameworks', 'Node.js', 'Active'],
      ['SKL-301', 'DevOps', 'Docker', 'Active'],
      ['SKL-302', 'DevOps', 'Kubernetes', 'Active'],
      ['SKL-401', 'Cloud', 'AWS', 'Active'],
      ['SKL-402', 'Cloud', 'GCP', 'Active'],
      ['SKL-501', 'AI', 'OpenAI', 'Active'],
      ['SKL-502', 'AI', 'LangChain', 'Active']
    ];
    defaultSkills.forEach(row => skillsSheet.appendRow(row));
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
      if (sheetName === SHEETS.ATT && header === 'Date') {
        obj[header] = value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(value);
        return;
      }
      if (header === 'CheckIn' || header === 'CheckOut') {
        obj[header] = value instanceof Date ? String(value.getTime()) : value || '';
        return;
      }
      if (sheetName === SHEETS.SET && row[0] === 'LateArrivalTime' && header === 'Value') {
        obj[header] = value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm") : String(value).trim();
        return;
      }
      if (value instanceof Date) {
        obj[header] = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
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

function updateRowByKey(sheetName, keyField, keyValue, updateData) {
  const sheet = getDb().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyColIndex = headers.indexOf(keyField);
  
  if (keyColIndex === -1) throw new Error(`Key field "${keyField}" not found.`);

  const rowIndex = data.findIndex(row => row[keyColIndex] == keyValue);
  if (rowIndex > 0) {
    for (const [updateKey, updateValue] of Object.entries(updateData)) {
      const colIndex = headers.indexOf(updateKey);
      if (colIndex > -1) sheet.getRange(rowIndex + 1, colIndex + 1).setValue(updateValue);
    }
    return true;
  }
  return false;
}

function logAudit(userEmail, action, details) {
  const sheet = getDb().getSheetByName(SHEETS.AUDIT);
  sheet.appendRow([Utilities.getUuid(), new Date().toISOString(), userEmail, action, details]);
}

function isDateHoliday(date, holidays) {
  if (!date || !(date instanceof Date)) return null;
  if (date.getDay() === 0) return { Name: 'Weekend', Type: 'Weekend', Description: 'Sunday Weekend' };
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return holidays.find(h => h.Date === dateStr) || null;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

// ==========================================
// SETTINGS
// ==========================================

function getAppSettings() {
  const data = getSheetDataAsJSON(SHEETS.SET);
  let settings = { DailyWorkingHours: 8, LateArrivalTime: "09:15", CasualLeaveQuota: 10, OptionalLeaveQuota: 2 };

  data.forEach(row => {
    const key = String(row.Key).trim();
    if (key === 'DailyWorkingHours') settings.DailyWorkingHours = Number(row.Value);
    if (key === 'CasualLeaveQuota') settings.CasualLeaveQuota = Number(row.Value);
    if (key === 'OptionalLeaveQuota') settings.OptionalLeaveQuota = Number(row.Value);
    if (key === 'LateArrivalTime') {
      settings.LateArrivalTime = row.Value instanceof Date ? Utilities.formatDate(row.Value, Session.getScriptTimeZone(), 'HH:mm') : String(row.Value).trim();
    }
  });
  return settings;
}

function saveAppSettings(settingsData) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  saveRowToSheet(SHEETS.SET, { Key: 'DailyWorkingHours', Value: settingsData.DailyWorkingHours }, 'Key');
  saveRowToSheet(SHEETS.SET, { Key: 'LateArrivalTime', Value: settingsData.LateArrivalTime }, 'Key');
  saveRowToSheet(SHEETS.SET, { Key: 'CasualLeaveQuota', Value: settingsData.CasualLeaveQuota }, 'Key');
  saveRowToSheet(SHEETS.SET, { Key: 'OptionalLeaveQuota', Value: settingsData.OptionalLeaveQuota }, 'Key');
  
  logAudit(currentUser, 'UPDATE_SETTINGS', `Updated System Settings & Quotas`);
  return { status: 'Success', message: 'Settings updated successfully!' };
}

// ==========================================
// AUTHENTICATION
// ==========================================

function loginUser(loginId, password) {
  if (!loginId || !password) return { status: 'Error', message: 'Login ID and password are required.' };

  const employees = getSheetDataAsJSON(SHEETS.EMP);
  const credential = loginId.trim().toLowerCase();
  const user = employees.find(e => (e.Email && e.Email.toLowerCase() === credential) || String(e.EmpID).toLowerCase() === credential);

  if (!user) return { status: 'Error', message: 'Account not found.' };
  if (user.Password !== password) return { status: 'Error', message: 'Invalid credentials.' };
  if (user.Status !== 'Active') return { status: 'Error', message: `Your account is inactive. Status: ${user.Status}.` };

  if (user.MustChangePassword === true || user.MustChangePassword === 'TRUE') {
    delete user.Password;
    return { status: 'PASSWORD_CHANGE_REQUIRED', user: user };
  }

  delete user.Password;
  logAudit(user.Email, 'LOGIN', `User logged in using ${credential}`);
  return { status: 'Success', user: user };
}

// ==========================================
// DASHBOARD
// ==========================================

function getDashboardData(currentUserEmpID, currentUserRole) {
  const todayDate = new Date();
  const todayStr = Utilities.formatDate(todayDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const emps = getSheetDataAsJSON(SHEETS.EMP);
  const atts = getSheetDataAsJSON(SHEETS.ATT);
  const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
  const hols = getSheetDataAsJSON(SHEETS.HOL);

  const todayHoliday = isDateHoliday(todayDate, hols);
  const upcomingHolidays = hols
    .filter(h => h.Date >= todayStr)
    .sort((a, b) => a.Date > b.Date ? 1 : -1)
    .map(h => {
      const holDate = new Date(h.Date);
      const todayMidnight = new Date(todayStr);
      const holMidnight = new Date(holDate.getFullYear(), holDate.getMonth(), holDate.getDate());
      h.daysLeft = Math.round((holMidnight.getTime() - todayMidnight.getTime()) / (1000 * 60 * 60 * 24));
      return h;
    });

  const nextHoliday = upcomingHolidays.find(h => h.Type !== 'Weekend') || null;

  const activeEmployees = emps.filter(e => e.Status === 'Active' && e.Role === 'EMPLOYEE');
  const activeEmpIds = activeEmployees.map(e => e.EmpID);

  const presentSet = new Set();
  atts.forEach(a => {
    if (a.CheckIn && activeEmpIds.includes(a.EmpID)) {
      if (String(a.Date).split('T')[0] === todayStr) presentSet.add(a.EmpID);
    }
  });

  const leavesTodayCount = leaves.filter(l => l.Status === 'Approved' && l.StartDate <= todayStr && l.EndDate >= todayStr && activeEmpIds.includes(l.EmpID)).length;

  const presentToday = todayHoliday ? 0 : presentSet.size;
  const absentToday = todayHoliday ? 0 : (activeEmployees.length - presentToday - leavesTodayCount);

  const trendLabels = [];
  const trendData = [];

  for (let i = 4; i >= 0; i--) {
    const d = new Date(todayDate.getTime() - (i * 24 * 60 * 60 * 1000));
    const dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    trendLabels.push(Utilities.formatDate(d, Session.getScriptTimeZone(), "EEE"));

    if (currentUserRole === 'ADMIN') {
      const presentOnDate = new Set();
      atts.forEach(a => {
        if (a.CheckIn && activeEmpIds.includes(a.EmpID) && String(a.Date).split('T')[0] === dStr) presentOnDate.add(a.EmpID);
      });
      trendData.push(activeEmployees.length > 0 ? Math.round((presentOnDate.size / activeEmployees.length) * 100) : 0);
    } else {
      let dailyHours = 0;
      atts.forEach(a => {
        if (a.EmpID === currentUserEmpID && String(a.Date).split('T')[0] === dStr && a.Hours) dailyHours += Number(a.Hours);
      });
      trendData.push(Number(dailyHours.toFixed(2)));
    }
  }

  return {
    totalEmployees: activeEmployees.length,
    presentToday: presentToday,
    absentToday: absentToday < 0 ? 0 : absentToday,
    onLeaveToday: leavesTodayCount,
    pendingLeaves: leaves.filter(l => l.Status === 'Pending').length,
    announcements: getSheetDataAsJSON(SHEETS.ANN).filter(a => a.Status === 'Active').slice(-5).reverse(),
    todayHoliday: todayHoliday,
    nextHoliday: nextHoliday,
    upcomingHolidays: upcomingHolidays.slice(0, 3),
    attendanceTrend: { labels: trendLabels, data: trendData }
  };
}

// ==========================================
// EMPLOYEE APIs
// ==========================================

function getEmployees() { return getSheetDataAsJSON(SHEETS.EMP).filter(e => e.Role === 'EMPLOYEE'); }

function generateEmpId() {
  const employees = getSheetDataAsJSON(SHEETS.EMP);
  let id;
  do { id = 'EMP-' + Math.floor(100000 + Math.random() * 900000); } while (employees.some(e => e.EmpID === id));
  return id;
}

function saveEmployee(empData) {
  const isNew = !empData.EmpID;
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  const tempPassword = 'Welcome@123';

  if (isNew) {
    empData.EmpID = generateEmpId();
    empData.JoiningDate = new Date().toISOString().split('T')[0];
    empData.Password = tempPassword;
    empData.MustChangePassword = true;
  }

  empData.Role = 'EMPLOYEE';
  if (!empData.Status) empData.Status = 'Active';

  saveRowToSheet(SHEETS.EMP, empData, 'EmpID');
  logAudit(currentUser, isNew ? 'CREATE_EMP' : 'UPDATE_EMP', `Employee ID: ${empData.EmpID}`);
  return { status: 'Success', message: isNew ? 'Employee added.' : 'Employee updated.', empId: empData.EmpID, password: isNew ? tempPassword : null };
}

function updateEmployeeStatus(empID, status) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  updateRowByKey(SHEETS.EMP, 'EmpID', empID, { Status: status, ExitDate: ['Resigned', 'Terminated', 'Retired'].includes(status) ? new Date() : '' });
  logAudit(currentUser, 'UPDATE_EMP_STATUS', `Employee ${empID} status to ${status}`);
  return { status: 'Success', message: 'Employee status updated.' };
}

function changePassword(loginId, currentPassword, newPassword) {
  const employees = getSheetDataAsJSON(SHEETS.EMP);
  const credential = String(loginId).trim().toLowerCase();
  const user = employees.find(e => (e.Email && e.Email.toLowerCase() === credential) || String(e.EmpID).toLowerCase() === credential);

  if (!user) return { status: 'Error', message: 'User not found.' };
  if (user.Password !== currentPassword) return { status: 'Error', message: 'Current password does not match.' };

  updateRowByKey(SHEETS.EMP, 'EmpID', user.EmpID, { Password: newPassword, MustChangePassword: false });
  return { status: 'Success', message: 'Password changed successfully.' };
}

function resetEmployeePassword(empId) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  const employees = getSheetDataAsJSON(SHEETS.EMP);
  const user = employees.find(e => String(e.EmpID).toLowerCase() === String(empId).trim().toLowerCase());

  if (!user) return { status: 'Error', message: 'Employee not found.' };

  updateRowByKey(SHEETS.EMP, 'EmpID', user.EmpID, { Password: 'Welcome@123', MustChangePassword: true });
  logAudit(currentUser, 'RESET_PASSWORD', `Password reset for ${user.EmpID}`);
  return { status: 'Success', message: `Password reset successfully.`, empId: user.EmpID, password: 'Welcome@123' };
}

function updateMyProfile(empId, profileData) {
  const currentUser = Session.getActiveUser().getEmail() || empId;
  const employees = getSheetDataAsJSON(SHEETS.EMP);
  const user = employees.find(e => e.EmpID === empId);
  
  if (!user) return { status: 'Error', message: 'Employee not found.' };

  if (profileData.Email && profileData.Email.toLowerCase() !== user.Email.toLowerCase()) {
    if (employees.find(e => e.Email.toLowerCase() === profileData.Email.toLowerCase())) {
      return { status: 'Error', message: 'Email address already in use.' };
    }
  }

  updateRowByKey(SHEETS.EMP, 'EmpID', empId, { Name: profileData.Name, Email: profileData.Email });
  logAudit(currentUser, 'UPDATE_PROFILE', `User ${empId} updated profile.`);
  return { status: 'Success', message: 'Profile updated successfully!' };
}

// ==========================================
// ATTENDANCE APIs
// ==========================================

function getAttendanceSummary(empIdFilter = null) {
  const rawData = getSheetDataAsJSON(SHEETS.ATT);
  const settings = getAppSettings();
  const reqHours = Number(settings.DailyWorkingHours) || 8;

  let filtered = empIdFilter ? rawData.filter(r => r.EmpID === empIdFilter) : rawData;
  const groups = {};

  filtered.forEach(row => {
    const key = row.EmpID + "_" + row.Date;
    if (!groups[key]) groups[key] = { EmpID: row.EmpID, Date: row.Date, TotalHours: 0, Sessions: [], RequiredHours: reqHours };
    groups[key].Sessions.push(row);
    if (!isNaN(Number(row.Hours))) groups[key].TotalHours += Number(row.Hours);
  });

  const result = Object.values(groups).map(g => {
    const hasActiveSession = g.Sessions.some(s => !s.CheckOut || s.CheckOut === "");
    if (hasActiveSession) g.Status = 'Present';
    else if (g.TotalHours >= g.RequiredHours) g.Status = 'Completed';
    else if (g.TotalHours >= (g.RequiredHours / 2)) g.Status = 'Half Day';
    else if (g.TotalHours > 0) g.Status = 'Present';
    else g.Status = 'Absent';

    g.Sessions.sort((a, b) => Number(String(a.CheckIn).replace(/,/g, '')) - Number(String(b.CheckIn).replace(/,/g, '')));
    return g;
  });

  result.sort((a, b) => a.Date > b.Date ? -1 : 1);
  return result;
}

function markAttendance(action, empId) {
  const now = new Date();
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const epochTimestamp = now.getTime();
  const currentUser = Session.getActiveUser().getEmail() || empId;
  const settings = getAppSettings();
  const sheet = getDb().getSheetByName(SHEETS.ATT);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  let lastRecordIndex = -1, lastCheckIn = null, lastCheckOut = null, lastDateStr = null, sessionsToday = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][headers.indexOf('EmpID')] === empId) {
      let d = new Date(data[i][headers.indexOf('Date')]);
      if (!isNaN(d.getTime())) {
        let rStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
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
    if (sessionsToday === 0) {
      const [lateH, lateM] = (settings.LateArrivalTime || "09:15").split(':').map(Number);
      const currentH = now.getHours(), currentM = now.getMinutes();
      if (currentH > lateH || (currentH === lateH && currentM > lateM)) isLate = 'Yes';
    }

    sheet.appendRow(['ATT-' + epochTimestamp, empId, todayStr, String(epochTimestamp), "", "", 'Present', isLate]);
    logAudit(currentUser, 'CHECK_IN', `Checked in at Epoch ${epochTimestamp}`);
    return { status: 'Success', message: 'Successfully checked in!' };

  } else if (action === 'CHECK_OUT') {
    if (!hasActiveSessionToday) return { status: 'Error', message: 'No active Check-In found.' };

    const cInEpoch = Number(String(lastCheckIn).replace(/,/g, ''));
    let workedHours = "0.00";
    if (!isNaN(cInEpoch) && cInEpoch > 0) workedHours = (Math.abs(epochTimestamp - cInEpoch) / 36e5).toFixed(2);

    sheet.getRange(lastRecordIndex, headers.indexOf('CheckOut') + 1).setValue(String(epochTimestamp));
    sheet.getRange(lastRecordIndex, headers.indexOf('Hours') + 1).setValue(workedHours);
    logAudit(currentUser, 'CHECK_OUT', `Checked out. Session Hours: ${workedHours}`);

    return { status: 'Success', message: 'Successfully checked out!' };
  }
  return { status: 'Error', message: 'Invalid action payload.' };
}

// ==========================================
// LEAVE & HOLIDAY APIs
// ==========================================

function getLeaves() { return getSheetDataAsJSON(SHEETS.LEAVE); }

function applyLeave(leaveData) {
  const currentUser = Session.getActiveUser().getEmail() || leaveData.EmpID;
  
  if (leaveData.Type === 'Optional Leave') {
    const holidays = getSheetDataAsJSON(SHEETS.HOL);
    const isOptHol = holidays.find(h => h.Date === leaveData.Date && h.Type === 'Optional');
    if (!isOptHol) return { status: 'Error', message: 'The selected date is not an Optional Holiday.' };

    const month = leaveData.Date.substring(0, 7);
    const leaves = getSheetDataAsJSON(SHEETS.LEAVE);
    const existingOpt = leaves.find(l => l.EmpID === leaveData.EmpID && l.Type === 'Optional Leave' && l.StartDate.substring(0, 7) === month && l.Status !== 'Rejected');

    if (existingOpt) return { status: 'Error', message: 'You have already selected an Optional Leave for this month.' };
  }

  leaveData.LeaveID = 'LV-' + Math.floor(100000 + Math.random() * 900000);
  leaveData.Status = 'Pending';
  leaveData.StartDate = leaveData.Date;
  leaveData.EndDate = leaveData.Date;
  leaveData.Days = 1;
  saveRowToSheet(SHEETS.LEAVE, leaveData, 'LeaveID');
  
  logAudit(currentUser, 'APPLY_LEAVE', `Applied for ${leaveData.Type}`);
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
// ANNOUNCEMENTS & DOCUMENTS
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

function uploadEmployeeDocument(fileData, docInfo, requestingEmail) {
  const adminUser = getSheetDataAsJSON(SHEETS.EMP).find(e => e.Email.toLowerCase() === requestingEmail.toLowerCase());
  if (!adminUser || adminUser.Role !== 'ADMIN') throw new Error('Permission Denied.');

  const { base64, mimeType, fileName } = fileData;
  const { empId, docType, month, year } = docInfo;

  const rootFolders = DriveApp.getFoldersByName('HRMS Documents');
  const rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder('HRMS Documents');
  const empFolder = getOrCreateFolder(rootFolder, empId);

  const decoded = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const driveFile = empFolder.createFile(blob);

  const docId = 'DOC-' + Date.now();
  getDb().getSheetByName(SHEETS.DOCUMENTS).appendRow([docId, empId, fileName, docType, month || '', year || '', driveFile.getId(), new Date().toISOString(), requestingEmail]);
  logAudit(requestingEmail, 'UPLOAD_DOCUMENT', `Uploaded '${fileName}' for ${empId}.`);
  return { status: 'Success', message: 'Document uploaded successfully.' };
}

function getEmployeeDocuments(requestingEmail) {
  const currentUser = getSheetDataAsJSON(SHEETS.EMP).find(e => e.Email.toLowerCase() === requestingEmail.toLowerCase());
  if (!currentUser) throw new Error('Access Denied.');
  const allDocuments = getSheetDataAsJSON(SHEETS.DOCUMENTS);
  if (currentUser.Role === 'ADMIN') return allDocuments.sort((a, b) => new Date(b.UploadDate) - new Date(a.UploadDate));
  return allDocuments.filter(doc => doc.EmpID === currentUser.EmpID).sort((a, b) => new Date(b.UploadDate) - new Date(a.UploadDate));
}

function getDocumentForDownload(docId, requestingEmail) {
  const currentUser = getSheetDataAsJSON(SHEETS.EMP).find(e => e.Email.toLowerCase() === requestingEmail.toLowerCase());
  const docMeta = getSheetDataAsJSON(SHEETS.DOCUMENTS).find(d => d.DocID === docId);
  if (!docMeta) throw new Error('Document not found.');
  if (currentUser.Role !== 'ADMIN' && currentUser.EmpID !== docMeta.EmpID) throw new Error('Access Denied.');

  try {
    const file = DriveApp.getFileById(docMeta.DriveFileID);
    const blob = file.getBlob();
    logAudit(requestingEmail, 'DOWNLOAD_DOCUMENT', `User downloaded document ${docId}.`);
    return { base64: Utilities.base64Encode(blob.getBytes()), fileName: docMeta.FileName, mimeType: blob.getContentType() };
  } catch (e) { throw new Error('File not found in Google Drive.'); }
}

function deleteEmployeeDocument(docId, requestingEmail) {
  const currentUser = getSheetDataAsJSON(SHEETS.EMP).find(e => e.Email.toLowerCase() === requestingEmail.toLowerCase());
  if (!currentUser || currentUser.Role !== 'ADMIN') throw new Error('Permission Denied.');
  const docMeta = getSheetDataAsJSON(SHEETS.DOCUMENTS).find(d => d.DocID === docId);
  if (!docMeta) throw new Error('Document metadata not found.');

  try { DriveApp.getFileById(docMeta.DriveFileID).setTrashed(true); } catch (e) {}
  deleteRowFromSheet(SHEETS.DOCUMENTS, 'DocID', docId);
  logAudit(requestingEmail, 'DELETE_DOCUMENT', `Deleted document ${docId}.`);
  return { status: 'Success', message: 'Document deleted successfully.' };
}

// ==========================================
// CLIENT APIs
// ==========================================

function getClients() { return getSheetDataAsJSON(SHEETS.CLIENTS); }
function getClientAssignments() { return getSheetDataAsJSON(SHEETS.CLIENT_ASSIGNMENTS); }

function saveClient(clientData) {
  const isNew = !clientData.ClientID;
  if (isNew) { clientData.ClientID = 'CL-' + Math.floor(100000 + Math.random() * 900000); clientData.CreatedAt = new Date().toISOString(); }
  saveRowToSheet(SHEETS.CLIENTS, clientData, 'ClientID');
  logAudit(Session.getActiveUser().getEmail() || 'Admin', isNew ? 'CREATE_CLIENT' : 'UPDATE_CLIENT', `Client: ${clientData.ClientName}`);
  return { status: 'Success', message: 'Client details saved successfully.' };
}

function deleteClient(clientId) {
  deleteRowFromSheet(SHEETS.CLIENTS, 'ClientID', clientId);
  const assignmentsSheet = getDb().getSheetByName(SHEETS.CLIENT_ASSIGNMENTS);
  const data = assignmentsSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) { if (data[i][1] === clientId) assignmentsSheet.deleteRow(i + 1); }
  logAudit(Session.getActiveUser().getEmail() || 'Admin', 'DELETE_CLIENT', `Removed Client ID: ${clientId}`);
  return { status: 'Success', message: 'Client deleted successfully.' };
}

function assignDevelopers(clientId, empIds) {
  const sheet = getDb().getSheetByName(SHEETS.CLIENT_ASSIGNMENTS);
  const data = sheet.getDataRange().getValues();
  const employees = getSheetDataAsJSON(SHEETS.EMP);
  for (let i = data.length - 1; i >= 1; i--) { if (data[i][1] === clientId) sheet.deleteRow(i + 1); }
  const today = new Date().toISOString().split('T')[0];
  empIds.forEach(empId => {
    const emp = employees.find(e => e.EmpID === empId);
    if (emp) sheet.appendRow(['ASN-' + Math.floor(100000 + Math.random() * 900000), clientId, empId, emp.Name, today, 'Active']);
  });
  logAudit(Session.getActiveUser().getEmail() || 'Admin', 'ASSIGN_DEVELOPERS', `Assigned devs to Client ${clientId}`);
  return { status: 'Success', message: 'Developers assigned successfully.' };
}

function getClientDashboardData(clientId) {
  const assignments = getSheetDataAsJSON(SHEETS.CLIENT_ASSIGNMENTS).filter(a => a.ClientID === clientId);
  const assignedEmpIds = assignments.map(a => a.EmpID);
  return {
    assignments,
    leaves: getSheetDataAsJSON(SHEETS.LEAVE).filter(l => assignedEmpIds.includes(l.EmpID)),
    attendance: getSheetDataAsJSON(SHEETS.ATT).filter(a => assignedEmpIds.includes(a.EmpID)),
    holidays: getSheetDataAsJSON(SHEETS.HOL)
  };
}

// ==========================================
// SKILLS MANAGEMENT APIs
// ==========================================

function getSkills() {
  const skills = getSheetDataAsJSON(SHEETS.SKILLS);
  const empSkills = getSheetDataAsJSON(SHEETS.EMP_SKILLS);
  
  const counts = {};
  empSkills.forEach(es => {
    counts[es.SkillID] = (counts[es.SkillID] || 0) + 1;
  });
  
  return skills.map(s => ({
    ...s,
    EmpCount: counts[s.SkillID] || 0
  }));
}

function saveSkill(skillData) {
  const currentUser = Session.getActiveUser().getEmail() || 'Admin';
  const skills = getSheetDataAsJSON(SHEETS.SKILLS);
  
  // Normalize the input name to prevent tricky duplicates (e.g. "React " vs "React")
  const normalizedNewName = skillData.SkillName.trim().toLowerCase();
  
  // Check if this skill name already exists (ignoring the one currently being edited)
  const duplicate = skills.find(s => 
    s.SkillName.trim().toLowerCase() === normalizedNewName && 
    s.SkillID !== skillData.SkillID
  );

  if (duplicate) {
    return { status: 'Error', message: `The skill '${skillData.SkillName}' already exists in the system.` };
  }

  const isNew = !skillData.SkillID;
  if (isNew) skillData.SkillID = 'SKL-' + Math.floor(10000 + Math.random() * 90000);
  
  saveRowToSheet(SHEETS.SKILLS, skillData, 'SkillID');
  logAudit(currentUser, 'SAVE_SKILL', `Saved skill: ${skillData.SkillName}`);
  
  return { status: 'Success', message: 'Skill successfully saved.' };
}

function deleteSkill(skillId) {
  deleteRowFromSheet(SHEETS.SKILLS, 'SkillID', skillId);
  const empSheet = getDb().getSheetByName(SHEETS.EMP_SKILLS);
  const data = empSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === skillId) empSheet.deleteRow(i + 1);
  }
  logAudit(Session.getActiveUser().getEmail() || 'Admin', 'DELETE_SKILL', `Deleted skill ID: ${skillId}`);
  return { status: 'Success', message: 'Skill and related employee records deleted.' };
}

function getEmployeeSkills(empId) {
  return getSheetDataAsJSON(SHEETS.EMP_SKILLS).filter(s => s.EmpID === empId);
}

function saveEmployeeSkills(empId, skillIds) {
  const sheet = getDb().getSheetByName(SHEETS.EMP_SKILLS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === empId) sheet.deleteRow(i + 1);
  }
  
  const masterSkills = getSheetDataAsJSON(SHEETS.SKILLS);
  const today = new Date().toISOString().split('T')[0];
  
  skillIds.forEach(id => {
    const sk = masterSkills.find(s => s.SkillID === id);
    if (sk && sk.Status === 'Active') {
      sheet.appendRow([empId, id, sk.SkillName, sk.Category, today]);
    }
  });
  
  logAudit(Session.getActiveUser().getEmail() || empId, 'UPDATE_EMP_SKILLS', `Updated skills for ${empId}`);
  return { status: 'Success', message: 'Your skills have been updated successfully!' };
}

function getEmployeesBySkill(skillId) {
  const empSkills = getSheetDataAsJSON(SHEETS.EMP_SKILLS).filter(es => es.SkillID === skillId);
  const empIds = empSkills.map(es => es.EmpID);
  
  if (empIds.length === 0) return [];
  
  const employees = getSheetDataAsJSON(SHEETS.EMP).filter(e => empIds.includes(e.EmpID));
  
  return employees.map(e => ({
    Name: e.Name,
    Designation: e.Designation,
    Department: e.Department
  }));
}