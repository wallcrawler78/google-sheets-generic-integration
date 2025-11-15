/**
 * PTC Arena Sheets Data Center
 * Main entry point for Google Sheets Add-on
 */

/**
 * Runs when the add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Runs when the spreadsheet is opened
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Arena Data Center')
    .addItem('Configure Arena Connection', 'showLoginWizard')
    .addItem('Test Connection', 'testArenaConnection')
    .addSeparator()
    .addItem('Import Data', 'importArenaData')
    .addItem('Clear Credentials', 'clearCredentials')
    .addToUi();
}

/**
 * Shows the login wizard for Arena API configuration
 */
function showLoginWizard() {
  var html = HtmlService.createHtmlOutputFromFile('LoginWizard')
    .setWidth(400)
    .setHeight(500)
    .setTitle('Configure Arena API Connection');
  SpreadsheetApp.getUi().showModalDialog(html, 'Arena API Configuration');
}

/**
 * Tests the Arena API connection
 */
function testArenaConnection() {
  var ui = SpreadsheetApp.getUi();

  if (!isAuthorized()) {
    ui.alert('Not Configured', 'Please configure your Arena API connection first.', ui.ButtonSet.OK);
    showLoginWizard();
    return;
  }

  try {
    var arenaClient = new ArenaAPIClient();
    var result = arenaClient.testConnection();

    if (result.success) {
      ui.alert('Success', 'Connection to Arena API successful!\n\nWorkspace: ' + getWorkspaceId(), ui.ButtonSet.OK);
    } else {
      ui.alert('Connection Failed', 'Could not connect to Arena API:\n' + result.error, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Error', 'Connection test failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Imports data from Arena (placeholder function)
 */
function importArenaData() {
  var ui = SpreadsheetApp.getUi();

  if (!isAuthorized()) {
    ui.alert('Not Configured', 'Please configure your Arena API connection first.', ui.ButtonSet.OK);
    showLoginWizard();
    return;
  }

  ui.alert('Import Data', 'Data import functionality coming soon!', ui.ButtonSet.OK);
}

/**
 * Clears stored credentials
 */
function clearCredentials() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Clear Credentials',
    'Are you sure you want to clear your Arena API credentials?',
    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    clearArenaCredentials();
    ui.alert('Cleared', 'Arena API credentials have been cleared.', ui.ButtonSet.OK);
  }
}
