/**
 * @fileoverview This file acts as the main entry point for the web app.
 */

/**
 * The main router for all GET requests to the web app.
 * It checks for parameters and routes the request to the appropriate handler function.
 * @param {Object} e The event parameter for a web app GET request.
 * @returns {HtmlOutput | ContentOutput} The output to be rendered in the browser.
 */
function doGet(e) {
  Logger.log('Master doGet triggered. Parameters: ' + JSON.stringify(e));
  
  if (e && e.parameter && e.parameter.action) {
    Logger.log('Routing to handleBirthdayApproval...');
    return handleBirthdayApproval(e);
  } else {
    Logger.log('Routing to showRecruitmentForm...');
    return showRecruitmentForm();
  }
}
