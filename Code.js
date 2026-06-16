/**
 * @fileoverview Main server file for the web app. Handles routing and serves HTML pages.
 */

/**
 * The main entry point for the web app. Called when a user accesses the URL.
 * @param {object} e - The event parameter containing request details.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML page to be served.
 */
function doGet(e) {
  const user = Auth.getAuthenticatedUser();

  // If user is not found in the database or not logged into Google, show login page.
  if (!user) {
    Utils.logAudit(Auth.getActiveUserEmail() || 'unknown', 'LOGIN_FAIL', 'User not found in database or not logged in.');
    return HtmlService.createTemplateFromFile('Login').evaluate()
      .setTitle('HRMS Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  }

  // Log successful login attempt
  if (e.parameter.page === undefined) { // Log only on initial load, not for sub-page navigation
      Utils.logAudit(user.Email, 'LOGIN_SUCCESS', `Role: ${user.Role}`);
  }

  const route = e.parameter.page || 'Dashboard'; // Default to Dashboard
  let template;

  // Basic routing based on the 'page' query parameter
  switch (route) {
    case 'Dashboard':
      template = HtmlService.createTemplateFromFile('Dashboard');
      break;
    case 'Attendance':
      template = HtmlService.createTemplateFromFile('Attendance');
      break;
    case 'Leave':
      template = HtmlService.createTemplateFromFile('Leave');
      break;
    // Add more routes here as we build more pages
    default:
      template = HtmlService.createTemplateFromFile('Dashboard'); // Fallback to Dashboard
  }

  // Pass user data to the template
  template.user = user;
  template.CONFIG = CONFIG; // Pass configuration to the frontend if needed

  return template.evaluate()
    .setTitle(`${CONFIG.COMPANY_NAME} HRMS`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Includes the content of another HTML file into the current template.
 * This is used for creating reusable UI components like sidebars and navbars.
 * @param {string} filename - The name of the HTML file to include (without extension).
 * @returns {string} The raw HTML content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * A server-side function to expose authenticated user data to the client-side JS.
 * @returns {object} The authenticated user's data.
 */
function getSessionUser() {
  return Auth.getAuthenticatedUser();
}
