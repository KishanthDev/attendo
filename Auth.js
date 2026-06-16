/**
 * @fileoverview Handles user authentication and role-based access control.
 */

const Auth = {
  /**
   * Gets the email of the currently logged-in Google user.
   * @returns {string} The user's email address.
   */
  getActiveUserEmail: function() {
    return Session.getActiveUser().getEmail();
  },

  /**
   * Retrieves user data (including role) from the Employees sheet.
   * @param {string} email - The email of the user to look up.
   * @returns {object|null} An object containing user data or null if not found.
   */
  getUserDataByEmail: function(email) {
    try {
      const ss = Utils.getSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.SHEETS.EMPLOYEES);
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();
      const emailCol = headers.indexOf("Email");
      const statusCol = headers.indexOf("Status");

      for (const row of data) {
        if (row[emailCol] === email && row[statusCol] === 'Active') {
          const userObject = {};
          headers.forEach((header, index) => {
            userObject[header] = row[index];
          });
          return userObject;
        }
      }
      return null; // User not found or is not active
    } catch (e) {
      console.error(`Error fetching user data for ${email}: ${e.toString()}`);
      return null;
    }
  },

  /**
   * Gets the complete authenticated user object.
   * This is the primary function to get user info for the session.
   * @returns {object|null} The user object or null if not authenticated.
   */
  getAuthenticatedUser: function() {
    const email = this.getActiveUserEmail();
    if (!email) {
      return null;
    }
    return this.getUserDataByEmail(email);
  },

  /**
   * Checks if the current user has Admin role.
   * @returns {boolean} True if the user is an Admin.
   */
  isAdmin: function() {
    const user = this.getAuthenticatedUser();
    return user && user.Role === CONFIG.ROLES.ADMIN;
  },

  /**
   * Checks if the current user has Manager role.
   * @returns {boolean} True if the user is a Manager.
   */
  isManager: function() {
    const user = this.getAuthenticatedUser();
    return user && (user.Role === CONFIG.ROLES.MANAGER || user.Role === CONFIG.ROLES.ADMIN);
  }
};
