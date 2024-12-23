// /src/sheets/basic/lodging.gs

/**
 * Lodging comparison sheet management
 * Dependencies: config.gs, utilities.gs, dataTables.gs
 */
const LodgingModule = {
  /**
   * Creates or updates lodging comparison sheet
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  setupLodgingSheet(ss) {
    const sheet =
      ss.getSheetByName(CONFIG.sheetNames.lodging) ||
      ss.insertSheet(CONFIG.sheetNames.lodging);
    sheet.clear();

    this.setupLayout(sheet);
    this.addValidation(sheet);
    this.addFormulas(sheet);
    this.createCharts(sheet);
  },

  /**
   * Sets up sheet layout and formatting
   * @param {Sheet} sheet - Lodging sheet
   */
  setupLayout(sheet) {
    sheet.setColumnWidths(1, 8, 150);

    const headers = [
      ["Accommodation Comparison", "", "", "", "", "", "", ""],
      [
        "Option",
        "Type",
        "Location",
        "Check-in",
        "Check-out",
        "Nightly Rate",
        "Total Cost",
        "Notes",
      ],
      [
        "Option 1",
        "",
        "",
        "",
        "",
        "",
        "=IF(AND(F3>0,E3>D3),F3*(E3-D3+1),0)",
        "",
      ],
      [
        "Option 2",
        "",
        "",
        "",
        "",
        "",
        "=IF(AND(F4>0,E4>D4),F4*(E4-D4+1),0)",
        "",
      ],
      [
        "Option 3",
        "",
        "",
        "",
        "",
        "",
        "=IF(AND(F5>0,E5>D5),F5*(E5-D5+1),0)",
        "",
      ],
    ];

    const range = sheet.getRange(1, 1, headers.length, 8);
    range.setValues(headers);

    UtilityModule.applyCellStyle(sheet.getRange("A1:H1"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A2:H2"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A3:H5"), "input");
  },

  /**
   * Adds data validation rules
   * @param {Sheet} sheet - Lodging sheet
   */
  addValidation(sheet) {
    // Accommodation type validation
    DataTablesModule.addReferenceValidation(
      sheet,
      "B3:B5",
      "accommodationTypes"
    );

    // Date validation
    UtilityModule.addDataValidation(sheet, "D3:E5", "date");

    // Cost validation
    UtilityModule.addDataValidation(sheet, "F3:F5", "number", {
      min: CONFIG.validation.budget.minAmount,
      max: CONFIG.validation.budget.maxAmount,
    });

    // Add date range validation
    for (let row = 3; row <= 5; row++) {
      UtilityModule.addDataValidation(sheet, `E${row}`, "custom", {
        formula: `=AND(E${row}>=D${row}, (E${row}-D${row}+1)>=${CONFIG.validation.trip.minDuration})`,
      });
    }
  },

  /**
   * Adds calculated fields
   * @param {Sheet} sheet - Lodging sheet
   */
  addFormulas(sheet) {
    const formulaRanges = ["G3:G5"];
    formulaRanges.forEach((range) => {
      UtilityModule.applyCellStyle(sheet.getRange(range), "formula");
    });
  },

  /**
   * Creates comparison charts
   * @param {Sheet} sheet - Lodging sheet
   */
  createCharts(sheet) {
    // Cost comparison chart
    const costChart = sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange("A3:A5"))
      .addRange(sheet.getRange("G3:G5"))
      .setPosition(7, 2, 0, 0)
      .setOption("title", "Total Cost Comparison")
      .setOption("width", CONFIG.charts.bar.width)
      .setOption("height", CONFIG.charts.bar.height)
      .setOption("colors", [CONFIG.styles.colors.accentColor])
      .build();

    sheet.insertChart(costChart);
  },
};
