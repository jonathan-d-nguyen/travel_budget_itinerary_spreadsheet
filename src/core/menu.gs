// /src/core/menu.gs

/**
 * Menu management for Travel Budget Planner
 * Dependencies: config.gs
 */
const MenuModule = {
  /**
   * Creates custom menu on spreadsheet open
   */
  onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Travel Planner")
      .addItem("Setup New Template", "setupTemplate")
      .addSeparator()
      .addSubMenu(this.createFileMenu())
      .addSubMenu(this.createToolsMenu())
      .addSubMenu(this.createHelpMenu())
      .addToUi();
  },

  /**
   * Creates File submenu items
   * @returns {Menu} Google Sheets Menu object
   */
  createFileMenu() {
    return SpreadsheetApp.getUi()
      .createMenu("File")
      .addItem("Reset All Sheets", "resetAllSheets")
      .addItem("Import Data", "importData")
      .addItem("Export Data", "exportData")
      .addSeparator()
      .addItem("Print Dashboard", "printDashboard");
  },

  /**
   * Creates Tools submenu items
   * @returns {Menu} Google Sheets Menu object
   */
  createToolsMenu() {
    return SpreadsheetApp.getUi()
      .createMenu("Tools")
      .addItem("Currency Converter", "showCurrencyConverter")
      .addItem("Date Calculator", "showDateCalculator")
      .addItem("Budget Optimizer", "showBudgetOptimizer")
      .addSeparator()
      .addItem("Preferences", "showPreferences");
  },

  /**
   * Creates Help submenu items
   * @returns {Menu} Google Sheets Menu object
   */
  createHelpMenu() {
    return SpreadsheetApp.getUi()
      .createMenu("Help")
      .addItem("Show Instructions", "showInstructions")
      .addItem("Load Example Data", "loadExampleData")
      .addSeparator()
      .addItem("About", "showAbout");
  },

  /**
   * Shows modal dialog with instructions
   */
  showInstructions() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const instructionsSheet = ss.getSheetByName(CONFIG.sheetNames.instructions);
    instructionsSheet.activate();
  },

  /**
   * Shows about dialog with version info
   */
  showAbout() {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Travel Budget Planner",
      `Version: ${CONFIG.version}\n\nA comprehensive tool for planning and tracking travel expenses.`,
      ui.ButtonSet.OK
    );
  },
};
