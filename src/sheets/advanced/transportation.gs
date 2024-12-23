// /src/sheets/advanced/transportation.gs

/**
 * Transportation planning sheet management
 * Dependencies: config.gs, utilities.gs, dataTables.gs
 */
const TransportationModule = {
  /**
   * Creates or updates transportation sheet
   * @param {Spreadsheet} ss - Active spreadsheet
   */
  setupTransportationSheet(ss) {
    const sheet =
      ss.getSheetByName(CONFIG.sheetNames.transportation) ||
      ss.insertSheet(CONFIG.sheetNames.transportation);
    sheet.clear();

    this.setupLayout(sheet);
    this.addValidation(sheet);
    this.addFormulas(sheet);
    this.createCharts(sheet);
  },

  /**
   * Sets up sheet layout and formatting
   * @param {Sheet} sheet - Transportation sheet
   */
  setupLayout(sheet) {
    sheet.setColumnWidths(1, 9, 120);

    const headers = [
      ["Transportation Planning", "", "", "", "", "", "", "", ""],
      [
        "Route",
        "Mode",
        "From",
        "To",
        "Date",
        "Duration (hrs)",
        "Cost",
        "Booking Ref",
        "Notes",
      ],
      ["", "", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", ""],
      ["Summary", "", "", "", "", "", "", "", ""],
      ["Total Duration (hrs)", "=SUM(F3:F7)", "", "", "", "", "", "", ""],
      ["Total Cost", "=SUM(G3:G7)", "", "", "", "", "", "", ""],
    ];

    const range = sheet.getRange(1, 1, headers.length, 9);
    range.setValues(headers);

    UtilityModule.applyCellStyle(sheet.getRange("A1:I1"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A2:I2"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A3:I7"), "input");
    UtilityModule.applyCellStyle(sheet.getRange("A8:I8"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A9:B10"), "formula");
  },

  /**
   * Adds data validation rules
   * @param {Sheet} sheet - Transportation sheet
   */
  addValidation(sheet) {
    // Transport mode validation
    DataTablesModule.addReferenceValidation(
      sheet,
      "B3:B7",
      "transportationModes"
    );

    // Date validation
    UtilityModule.addDataValidation(sheet, "E3:E7", "date");

    // Duration validation
    UtilityModule.addDataValidation(sheet, "F3:F7", "number", {
      min: 0,
      max: 24,
    });

    // Cost validation
    UtilityModule.addDataValidation(sheet, "G3:G7", "number", {
      min: CONFIG.validation.budget.minAmount,
      max: CONFIG.validation.budget.maxAmount,
    });
  },

  /**
   * Adds calculated fields and summary formulas
   * @param {Sheet} sheet - Transportation sheet
   */
  addFormulas(sheet) {
    UtilityModule.applyCellStyle(sheet.getRange("B9:B10"), "formula");
  },

  /**
   * Creates visualization charts
   * @param {Sheet} sheet - Transportation sheet
   */
  createCharts(sheet) {
    // Mode breakdown pie chart
    const modeChart = sheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("B3:B7"))
      .addRange(sheet.getRange("G3:G7"))
      .setPosition(12, 2, 0, 0)
      .setOption("title", "Cost by Transport Mode")
      .setOption("width", CONFIG.charts.pie.width)
      .setOption("height", CONFIG.charts.pie.height)
      .build();

    sheet.insertChart(modeChart);

    // Timeline chart
    const timelineChart = sheet
      .newChart()
      .setChartType(Charts.ChartType.TIMELINE)
      .addRange(sheet.getRange("A3:F7"))
      .setPosition(12, 6, 0, 0)
      .setOption("title", "Transportation Timeline")
      .setOption("width", CONFIG.charts.timeline.width)
      .setOption("height", CONFIG.charts.timeline.height)
      .build();

    sheet.insertChart(timelineChart);
  },
};
