// /src/sheets/basic/tripSettings.gs

/**
 * Trip Settings sheet management
 * Dependencies: config.gs, utilities.gs, dataTables.gs
 */
const TripSettingsModule = {
  /**
   * Creates or updates trip settings sheet
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  setupTripSettingsSheet(ss) {
    const sheet =
      ss.getSheetByName(CONFIG.sheetNames.tripSettings) ||
      ss.insertSheet(CONFIG.sheetNames.tripSettings);
    sheet.clear();

    this.setupLayout(sheet);
    this.addValidation(sheet);
    this.addFormulas(sheet);
    this.protectSheet(sheet);
  },

  /**
   * Sets up sheet layout and formatting
   * @param {Sheet} sheet - Trip Settings sheet
   */
  setupLayout(sheet) {
    // Set column widths
    sheet.setColumnWidths(1, 2, 200);
    sheet.setColumnWidth(3, 300);

    const headers = [
      ["Trip Details", "", ""],
      ["Trip Name", "", ""],
      ["Trip Type", "", ""],
      ["Start Date", "", ""],
      ["End Date", "", ""],
      ["Number of Travelers", "", ""],
      ["", "", ""],
      ["Budget Settings", "", ""],
      ["Total Budget", "", ""],
      ["Currency", "", ""],
      ["Budget per Person", "=IF(E6>0, E9/E6, 0)", ""],
      ["Daily Budget", "=IF(E15>0, E9/E15, 0)", ""],
      ["", "", ""],
      ["Calculations", "", ""],
      ["Total Days", '=IF(AND(E4<>"",E5<>""), E5-E4+1, 0)', ""],
    ];

    const range = sheet.getRange(1, 1, headers.length, 3);
    range.setValues(headers);

    // Apply styles
    headers.forEach((row, index) => {
      if (
        row[0].endsWith("Settings") ||
        row[0] === "Trip Details" ||
        row[0] === "Calculations"
      ) {
        UtilityModule.applyCellStyle(sheet.getRange(index + 1, 1), "header");
      } else if (row[0]) {
        UtilityModule.applyCellStyle(
          sheet.getRange(index + 1, 1, 1, 2),
          "input"
        );
      }
    });
  },

  /**
   * Adds data validation rules
   * @param {Sheet} sheet - Trip Settings sheet
   */
  addValidation(sheet) {
    // Trip type validation
    DataTablesModule.addReferenceValidation(sheet, "B3", "tripTypes");

    // Date validation
    UtilityModule.addDataValidation(sheet, "B4:B5", "date");

    // Numeric validation
    UtilityModule.addDataValidation(sheet, "B6", "number", {
      min: CONFIG.validation.trip.minTravelers,
      max: CONFIG.validation.trip.maxTravelers,
    });

    UtilityModule.addDataValidation(sheet, "B9", "number", {
      min: CONFIG.validation.budget.minAmount,
      max: CONFIG.validation.budget.maxAmount,
    });

    // Custom date range validation
    UtilityModule.addDataValidation(sheet, "B5", "custom", {
      formula: `=AND(B5>=B4, (B5-B4+1)>=${CONFIG.validation.trip.minDuration}, (B5-B4+1)<=${CONFIG.validation.trip.maxDuration})`,
    });
  },

  /**
   * Adds calculated fields
   * @param {Sheet} sheet - Trip Settings sheet
   */
  addFormulas(sheet) {
    const formulaRanges = ["B11:B12", "B15"];
    formulaRanges.forEach((range) => {
      UtilityModule.applyCellStyle(sheet.getRange(range), "formula");
    });
  },

  /**
   * Protects sheet and formula cells
   * @param {Sheet} sheet - Trip Settings sheet
   */
  protectSheet(sheet) {
    ProtectionModule.protectFormulas(sheet, "B11:B15");
  },
};
