/**
 * @fileoverview Configuration file for the HRMS application.
 * Contains global constants and settings.
 */

// Main configuration object
const CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  SHEETS: {
    EMPLOYEES: "Employees",
    ATTENDANCE: "Attendance",
    LEAVE_REQUESTS: "Leave Requests",
    HOLIDAYS: "Holidays",
    ANNOUNCEMENTS: "Announcements",
    TIMESHEET: "Timesheet",
    AUDIT_LOGS: "AuditLogs"
  },
  ROLES: {
    ADMIN: "Admin",
    MANAGER: "Manager",
    EMPLOYEE: "Employee"
  },
  LEAVE_STATUS: {
    PENDING: "Pending",
    APPROVED: "Approved",
    REJECTED: "Rejected",
    CANCELLED: "Cancelled"
  },
  ATTENDANCE_STATUS: {
    PRESENT: "Present",
    ABSENT: "Absent",
    HALF_DAY: "Half Day",
    LEAVE: "Leave",
    WFH: "Work From Home"
  },
  COMPANY_NAME: "Your Company Name"
};
