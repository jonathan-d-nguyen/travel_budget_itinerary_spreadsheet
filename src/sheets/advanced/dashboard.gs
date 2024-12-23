// /src/sheets/advanced/dashboard.gs

/**
 * Dashboard sheet management for trip overview
 * Dependencies: config.gs, utilities.gs
 */
const DashboardModule = {
  setupDashboardSheet(ss) {
    const sheet =
      ss.getSheetByName(CONFIG.sheetNames.dashboard) ||
      ss.insertSheet(CONFIG.sheetNames.dashboard);
    sheet.clear();

    this.setupLayout(sheet);
    this.addFormulas(sheet);
    this.createCharts(sheet);
  },

  setupLayout(sheet) {
    sheet.setColumnWidths(1, 6, 150);

    const headers = [
      ["Trip Overview Dashboard", "", "", "", "", ""],
      ["Trip Details", "", "", "", "", ""],
      [
        "Trip Name",
        "=INDIRECT(\"'Trip Settings'!B2\")",
        "",
        "Budget Overview",
        "",
        "",
      ],
      [
        "Duration",
        "=INDIRECT(\"'Trip Settings'!B15\")",
        "days",
        "Total Budget",
        "=INDIRECT(\"'Trip Settings'!B9\")",
        "",
      ],
      [
        "Start Date",
        "=INDIRECT(\"'Trip Settings'!B4\")",
        "",
        "Spent So Far",
        "=SUM(INDIRECT(\"'Lodging'!G3:G5\"),INDIRECT(\"'Transportation'!B10\"),INDIRECT(\"'Activities'!B11\"))",
        "",
      ],
      [
        "End Date",
        "=INDIRECT(\"'Trip Settings'!B5\")",
        "",
        "Remaining",
        "=B4-B5",
        "",
      ],
      ["", "", "", "", "", ""],
      ["Expense Breakdown", "", "", "", "", ""],
      ["Category", "Amount", "% of Budget", "", "", ""],
      [
        "Lodging",
        "=INDIRECT(\"'Lodging'!B10\")",
        "=B10/INDIRECT(\"'Trip Settings'!B9\")",
        "",
        "",
        "",
      ],
      [
        "Transportation",
        "=INDIRECT(\"'Transportation'!B10\")",
        "=B11/INDIRECT(\"'Trip Settings'!B9\")",
        "",
        "",
        "",
      ],
      [
        "Activities",
        "=INDIRECT(\"'Activities'!B11\")",
        "=B12/INDIRECT(\"'Trip Settings'!B9\")",
        "",
        "",
        "",
      ],
    ];

    const range = sheet.getRange(1, 1, headers.length, 6);
    range.setValues(headers);

    // Apply styles
    UtilityModule.applyCellStyle(sheet.getRange("A1:F1"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A2:B2"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("D2:E2"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("A8:C8"), "header");
    UtilityModule.applyCellStyle(sheet.getRange("B3:B6"), "formula");
    UtilityModule.applyCellStyle(sheet.getRange("E4:E6"), "formula");
    UtilityModule.applyCellStyle(sheet.getRange("B10:C12"), "formula");
  },

  addFormulas(sheet) {
    // Format percentages
    sheet.getRange("C10:C12").setNumberFormat("0.0%");

    // Format currency
    sheet.getRange("B10:B12").setNumberFormat("$#,##0.00");
    sheet.getRange("E4:E6").setNumberFormat("$#,##0.00");

    // Format dates
    sheet.getRange("B4:B5").setNumberFormat("mm/dd/yyyy");
  },

  createCharts(sheet) {
    // Budget overview pie chart
    const budgetChart = sheet
      .newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange("A10:B12"))
      .setPosition(14, 1, 0, 0)
      .setOption("title", "Expense Distribution")
      .setOption("width", CONFIG.charts.pie.width)
      .setOption("height", CONFIG.charts.pie.height)
      .build();

    sheet.insertChart(budgetChart);

    // Budget vs Actual bar chart
    const comparisonChart = sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange("D4:E5"))
      .setPosition(14, 5, 0, 0)
      .setOption("title", "Budget vs Spent")
      .setOption("width", CONFIG.charts.bar.width)
      .setOption("height", CONFIG.charts.bar.height)
      .build();

    sheet.insertChart(comparisonChart);
  },
};
