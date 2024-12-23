// /src/sheets/basic/instructions.gs

/**
 * Instructions sheet management
 * Dependencies: config.gs, utilities.gs
 */
const InstructionsModule = {
  /**
   * Creates or updates the instructions sheet
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  setupInstructionsSheet(ss) {
    const sheet =
      ss.getSheetByName(CONFIG.sheetNames.instructions) ||
      ss.insertSheet(CONFIG.sheetNames.instructions);
    sheet.clear();

    // Set column widths
    sheet.setColumnWidth(1, 800);

    // Add sections
    this.addWelcomeSection(sheet);
    this.addGettingStartedSection(sheet);
    this.addSheetDescriptionsSection(sheet);
    this.addTipsSection(sheet);

    // Format and protect
    sheet.setFrozenRows(1);
    ProtectionModule.protectSheets(ss, [CONFIG.sheetNames.instructions]);
  },

  /**
   * Adds welcome section with overview
   * @param {Sheet} sheet - Instructions sheet
   */
  addWelcomeSection(sheet) {
    const range = sheet.getRange("A1:A3");
    range.setValues([
      ["Welcome to Travel Budget Planner"],
      ["Version " + CONFIG.version],
      ["This template helps you plan and track travel expenses efficiently."],
    ]);
    UtilityModule.applyCellStyle(sheet.getRange("A1"), "header");
  },

  /**
   * Adds getting started instructions
   * @param {Sheet} sheet - Instructions sheet
   */
  addGettingStartedSection(sheet) {
    const startRow = 5;
    const content = [
      ["Getting Started"],
      ["1. Fill in your trip details in the Trip Settings sheet"],
      ["2. Compare accommodation options in the Lodging sheet"],
      ["3. Plan transportation and activities"],
      ["4. Review your budget in the Dashboard"],
    ];

    const range = sheet.getRange(startRow, 1, content.length, 1);
    range.setValues(content);
    UtilityModule.applyCellStyle(sheet.getRange(startRow, 1), "header");
  },

  /**
   * Adds sheet descriptions
   * @param {Sheet} sheet - Instructions sheet
   */
  addSheetDescriptionsSection(sheet) {
    const startRow = 11;
    const content = [
      ["Sheet Descriptions"],
      ["Trip Settings: Basic trip information and preferences"],
      ["Lodging: Compare different accommodation options"],
      ["Transportation: Plan travel between locations"],
      ["Activities: Schedule and budget activities"],
      ["Dashboard: Overview of trip timeline and costs"],
    ];

    const range = sheet.getRange(startRow, 1, content.length, 1);
    range.setValues(content);
    UtilityModule.applyCellStyle(sheet.getRange(startRow, 1), "header");
  },

  /**
   * Adds usage tips and best practices
   * @param {Sheet} sheet - Instructions sheet
   */
  addTipsSection(sheet) {
    const startRow = 18;
    const content = [
      ["Tips and Best Practices"],
      ["• Enter data in white or light blue cells only"],
      ["• Grey cells contain formulas - do not edit directly"],
      ["• Use the Tools menu for currency conversion"],
      ["• Save regularly using File > Save"],
    ];

    const range = sheet.getRange(startRow, 1, content.length, 1);
    range.setValues(content);
    UtilityModule.applyCellStyle(sheet.getRange(startRow, 1), "header");
  },
};
