/**
 * @fileoverview This file contains the initial setup function to create and configure the Google Sheets database.
 * This function should be run manually once from the Apps Script editor to initialize the project.
 */

/**
 * Creates the main spreadsheet and all necessary sheets with headers.
 * Stores the new spreadsheet's ID in Script Properties for later use.
 *
 * @returns {void}
 */
function setupSpreadsheets() {
  // Create a new Spreadsheet with a specific name
  const spreadsheet = SpreadsheetApp.create("HRMS_Database");
  const spreadsheetId = spreadsheet.getId();

  // Store the ID in PropertiesService so the app can find it
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);

  // Log the URL for easy access
  console.log(`Spreadsheet created. URL: ${spreadsheet.getUrl()}`);

  // Get the default 'Sheet1' to delete it later
  const defaultSheet = spreadsheet.getSheetByName('Sheet1');

  // Define sheet names and their respective headers
  const sheets = {
    "Employees": [
      "EmployeeID", "Name", "Email", "Phone", "Department", "Designation",
      "Role", "ManagerID", "LeaveBalance", "Status", "DateCreated", "ProfilePictureUrl", "Address", "EmergencyContact"
    ],
    "Attendance": [
      "AttendanceID", "EmployeeID", "Date", "CheckIn", "CheckOut",
      "HoursWorked", "AttendanceStatus", "Location", "Remarks"
    ],
    "Leave Requests": [
      "LeaveID", "EmployeeID", "LeaveType", "StartDate", "EndDate",
      "Reason", "Status", "ManagerComment", "AdminComment", "CreatedDate"
    ],
    "Holidays": [
      "HolidayID", "HolidayName", "Date", "Description"
    ],
    "Announcements": [
      "AnnouncementID", "Title", "Description", "CreatedBy", "CreatedDate"
    ],
    "Timesheet": [
      "TimesheetID", "EmployeeID", "Date", "Project", "Task", "Hours", "Notes"
    ],
    "AuditLogs": [
      "LogID", "Timestamp", "UserEmail", "Action", "Details"
    ]
  };

  // Create each sheet and set its headers
  for (const sheetName in sheets) {
    const headers = sheets[sheetName];
    const sheet = spreadsheet.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    console.log(`'' sheet created with headers.`);
  }

  // Add a default admin user for initial login
  const employeesSheet = spreadsheet.getSheetByName("Employees");
  if (employeesSheet) {
    const adminEmail = Session.getActiveUser().getEmail();
    employeesSheet.appendRow([
      "EMP001", // EmployeeID
      "Admin User", // Name
      adminEmail, // Email
      "", // Phone
      "Management", // Department
      "System Administrator", // Designation
      "Admin", // Role
      "", // ManagerID
      12, // LeaveBalance
      "Active", // Status
      new Date(), // DateCreated
      "", "", "" // ProfilePictureUrl, Address, EmergencyContact
    ]);
    console.log(`Default admin user added with email: `);
  }


  // Delete the default 'Sheet1'
  if (defaultSheet) {
    spreadsheet.deleteSheet(defaultSheet);
  }

  // Set protections and validations if needed (can be expanded)
  SpreadsheetApp.flush(); // Apply all pending changes
}
