/**
 * @fileoverview Utility functions used throughout the application.
 */

const Utils = {
  /**
   * Gets the active Spreadsheet database instance.
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The spreadsheet object.
   * @throws {Error} If SPREADSHEET_ID is not configured.
   */
  getSpreadsheet: function() {
    if (!CONFIG.SPREADSHEET_ID) {
      throw new Error("Spreadsheet ID not found. Please run the 'setupSpreadsheets' function from Setup.gs first.");
    }
    try {
      return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } catch (e) {
      console.error("Failed to open spreadsheet. Check if the ID is correct and you have permissions.", e);
      throw new Error("Could not connect to the database. Please verify configuration and permissions.");
    }
  },

  /**
   * Generates a unique ID.
   * @param {string} prefix - A prefix for the ID (e.g., 'EMP', 'LEV').
   * @returns {string} A unique identifier string.
   */
  generateUniqueId: function(prefix) {
    return `${prefix}${Date.now()}${Math.floor(Math.random() * 1000)}`;
  },

  /**
   * Sends an email using Google's MailApp service.
   * @param {string} recipient - The email address of the recipient.
   * @param {string} subject - The subject of the email.
   * @param {string} body - The HTML body of the email.
   */
  sendEmail: function(recipient, subject, body) {
    try {
      MailApp.sendEmail({
        to: recipient,
        subject: subject,
        htmlBody: body
      });
    } catch (e) {
      console.error(`Failed to send email to : ${e.toString()}`);
    }
  },

  /**
   * Logs an action to the AuditLogs sheet for tracking purposes.
   * @param {string} userEmail - The email of the user performing the action.
   * @param {string} action - A description of the action (e.g., 'LOGIN_SUCCESS', 'LEAVE_APPROVED').
   * @param {string} details - Additional details about the action.
   */
  logAudit: function(userEmail, action, details) {
    try {
      const ss = this.getSpreadsheet();
      const auditSheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOGS);
      const logId = this.generateUniqueId('LOG');
      auditSheet.appendRow([logId, new Date(), userEmail, action, details]);
    } catch (e) {
      console.error(`Failed to write to audit log: ${e.toString()}`);
    }
  }
};
