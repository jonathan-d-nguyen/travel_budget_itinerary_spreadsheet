// /src/sheets/basic/dataTables.gs

/**
 * Data Tables management for reference data
 * Dependencies: config.gs, utilities.gs
 */
const DataTablesModule = {
  /**
   * Creates or updates the data tables sheet
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  setupDataTablesSheet(ss) {
    const sheet =
      ss.getSheetByName(CONFIG.sheetNames.dataTables) ||
      ss.insertSheet(CONFIG.sheetNames.dataTables);
    sheet.clear();

    this.createReferenceTable(
      sheet,
      "Trip Types",
      CONFIG.referenceData.tripTypes,
      "A"
    );
    this.createReferenceTable(
      sheet,
      "Accommodation",
      CONFIG.referenceData.accommodationTypes,
      "C"
    );
    this.createReferenceTable(
      sheet,
      "Transportation",
      CONFIG.referenceData.transportationModes,
      "E"
    );
    this.createReferenceTable(
      sheet,
      "Activities",
      CONFIG.referenceData.activityCategories,
      "G"
    );

    // Hide sheet and protect
    sheet.hideSheet();
    ProtectionModule.protectSheets(ss, [CONFIG.sheetNames.dataTables]);
  },

  /**
   * Creates a reference data table with named range
   * @param {Sheet} sheet - Data Tables sheet
   * @param {string} title - Table title
   * @param {string[]} data - Array of values
   * @param {string} column - Column letter for table
   */
  createReferenceTable(sheet, title, data, column) {
    const startRow = 1;
    const range = `${column}${startRow}:${column}${startRow + data.length}`;

    // Set title and data
    sheet.getRange(`${column}${startRow}`).setValue(title);
    UtilityModule.applyCellStyle(
      sheet.getRange(`${column}${startRow}`),
      "header"
    );

    // Set data values
    const dataRange = sheet.getRange(
      `${column}${startRow + 1}:${column}${startRow + data.length}`
    );
    dataRange.setValues(data.map((item) => [item]));
    UtilityModule.applyCellStyle(dataRange, "input");

    // Create named range
    const rangeName = title.replace(/\s+/g, "") + "List";
    UtilityModule.createNamedRange(
      sheet.getParent(),
      sheet,
      rangeName,
      `${column}${startRow + 1}:${column}${startRow + data.length}`
    );
  },

  /**
   * Gets reference data range name
   * @param {string} category - Category name
   * @returns {string} Named range reference
   */
  getReferenceRange(category) {
    return `=${category.replace(/\s+/g, "")}List`;
  },

  /**
   * Adds data validation to a range using reference data
   * @param {Sheet} sheet - Target sheet
   * @param {string} range - Range in A1 notation
   * @param {string} category - Reference data category
   */
  addReferenceValidation(sheet, range, category) {
    UtilityModule.addDataValidation(sheet, range, "list", {
      values: CONFIG.referenceData[category],
      strictValidation: true,
    });
  },
};
