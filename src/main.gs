// /src/main.gs

/**
 * Main entry point and initialization
 * Dependencies: All modules
 */

function onOpen() {
  MenuModule.onOpen();
}

/**
 * Sets up a new travel budget planner template
 */
function setupTemplate() {
  const ss = SpreadsheetApp.getActive();

  // Initialize sheets in order
  InstructionsModule.setupInstructionsSheet(ss);
  DataTablesModule.setupDataTablesSheet(ss);
  TripSettingsModule.setupTripSettingsSheet(ss);
  LodgingModule.setupLodgingSheet(ss);
  TransportationModule.setupTransportationSheet(ss);
  ActivitiesModule.setupActivitiesSheet(ss);
  DashboardModule.setupDashboardSheet(ss);

  // Final setup
  ss.setActiveSheet(ss.getSheetByName(CONFIG.sheetNames.instructions));
  SpreadsheetApp.flush();

  // Show confirmation
  SpreadsheetApp.getUi().alert(
    "Setup Complete",
    "The travel budget planner template has been created successfully.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Menu callback functions
function resetAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirm Reset",
    "This will clear all data. Are you sure you want to continue?",
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    setupTemplate();
  }
}

function showCurrencyConverter() {
  // TODO: Implement currency converter dialog
}

function showDateCalculator() {
  // TODO: Implement date calculator dialog
}

function showBudgetOptimizer() {
  // TODO: Implement budget optimizer dialog
}

function showPreferences() {
  // TODO: Implement preferences dialog
}

function loadExampleData() {
  // TODO: Implement example data loader
}

function importData() {
  // TODO: Implement data import functionality
}

function exportData() {
  // TODO: Implement data export functionality
}

function printDashboard() {
  const ss = SpreadsheetApp.getActive();
  const dashboard = ss.getSheetByName(CONFIG.sheetNames.dashboard);
  dashboard.activate();
  SpreadsheetApp.getUi().alert(
    "Print Dashboard",
    "Please use File > Print from the Google Sheets menu to print the dashboard.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
