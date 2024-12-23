// /src/sheets/advanced/activities.gs

/**
 * Activities planning sheet management
 * Dependencies: config.gs, utilities.gs, dataTables.gs
 */
const ActivitiesModule = {
  setupActivitiesSheet(ss) {
    const sheet =
      ss.getSheetByName(CONFIG.sheetNames.activities) ||
      ss.insertSheet(CONFIG.sheetNames.activities);
    sheet.clear();

    this.setupLayout(sheet);
    this.addValidation(sheet);
    this.addFormulas(sheet);
    this.createCharts(sheet);
  },

  setupLayout(sheet) {
    sheet.setColumnWidths(1, 10, 120);

    const headers = [
      ["Activities Planner", "", "", "", "", "", "", "", "", ""],
      [
        "Day",
        "Date",
        "Activity",
        "Category",
        "Start Time",
        "Duration (hrs)",
        "Cost per Person",
        "Total People",
        "Total Cost",
        "Notes",
      ],
      ["1", "", "", "", "", "", "", "", "=G3*H3", ""],
      ["2", "", "", "", "", "", "", "", "=G4*H4", ""],
      ["3", "", "", "", "", "", "", "", "=G5*H5", ""],
      ["4", "", "", "", "", "", "", "", "=G6*H6", ""],
      ["5", "", "", "", "", "", "", "", "=G7*H7", ""],
      ["Summary", "", "", "", "", "", "", "", "", ""],
      ["Total Activities", "=COUNTA(C3:C7)", "", "", "", "", "", "", "", ""],
      ["Total Duration", "=SUM(F3:F7)", "", "", "", "", "", "", "", ""],
      ["Total Cost", "=SUM(I3:I7)", "", "", "", "", "", "", "", ""],
    ];

    const range = sheet.getRange(1, 1, headers.length, 10);
    range.setValues(headers);

    // Apply styles
    UtilityModule.applyCellStyle(sheet.getRange("A1:J1"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A2:J2"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A3:J7"), "input");
    UtilityModule.applyCellStyle(sheet.getRange("A8:J8"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A9:B11"), "formula");
    UtilityModule.applyCellStyle(sheet.getRange("I3:I7"), "formula");
  },

  addValidation(sheet) {
    // Activity category validation
    DataTablesModule.addReferenceValidation(
      sheet,
      "D3:D7",
      "activityCategories"
    );

    // Date validation
    UtilityModule.addDataValidation(sheet, "B3:B7", "date");

    // Time validation
    UtilityModule.addDataValidation(sheet, "E3:E7", "custom", {
      formula: "=AND(E3>=0,E3<24)",
    });

    // Duration validation
    UtilityModule.addDataValidation(sheet, "F3:F7", "number", {
      min: CONFIG.validation.activities.minDuration,
      max: CONFIG.validation.activities.maxDuration,
    });

    // Cost validation
    UtilityModule.addDataValidation(sheet, "G3:G7", "number", {
      min: CONFIG.validation.budget.minAmount,
      max: CONFIG.validation.budget.maxAmount,
    });

    // People count validation
    UtilityModule.addDataValidation(sheet, "H3:H7", "number", {
      min: 1,
      max: CONFIG.validation.trip.maxTravelers,
    });
  },

  addFormulas(sheet) {
    // Protected formula cells
    ProtectionModule.protectFormulas(sheet, "I3:I7");
    ProtectionModule.protectFormulas(sheet, "B9:B11");
  },

  createCharts(sheet) {
    // Cost by category pie chart
    const costChart = sheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("D3:D7"))
      .addRange(sheet.getRange("I3:I7"))
      .setPosition(13, 2, 0, 0)
      .setOption("title", "Cost by Activity Category")
      .setOption("width", CONFIG.charts.pie.width)
      .setOption("height", CONFIG.charts.pie.height)
      .build();

    sheet.insertChart(costChart);

    // Daily timeline chart
    const timelineChart = sheet
      .newChart()
      .setChartType(Charts.ChartType.TIMELINE)
      .addRange(sheet.getRange("B3:F7"))
      .setPosition(13, 6, 0, 0)
      .setOption("title", "Activities Timeline")
      .setOption("width", CONFIG.charts.timeline.width)
      .setOption("height", CONFIG.charts.timeline.height)
      .build();

    sheet.insertChart(timelineChart);
  },
};
