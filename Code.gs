/**
 * Code.gs
 * Combined Google Apps Script modules for Travel Budget Planner
 * Generated on: 2024-12-23T06:14:01.731Z
 */


// ================================================================================
// Module: config.gs
// ================================================================================

// /src/config/config.gs

/**
 * Global configuration object for Travel Budget Planner
 */
const CONFIG = {
  version: "1.0.0",
  sheetNames: {
    instructions: "Instructions",
    tripSettings: "Trip Settings",
    lodging: "Lodging Comparison",
    transportation: "Transportation",
    activities: "Activities",
    dataTables: "DataTables",
    calculations: "Calculations",
    dashboard: "Dashboard",
    examples: "Example Data",
  },

  styles: {
    colors: {
      headerBg: "#1F4E79", // Dark blue headers
      headerText: "#FFFFFF", // White text
      inputBg: "#E6F3FF", // Light blue input cells
      formulaBg: "#F2F2F2", // Light grey formula cells
      warningBg: "#FFE6E6", // Light red warnings
      altRowBg: "#F8FBFF", // Alternate row background
      borderColor: "#CCCCCC", // Grey borders
      accentColor: "#4472C4", // Accent color for charts
    },
    fonts: {
      header: "Arial",
      body: "Arial",
    },
  },

  validation: {
    trip: {
      minDuration: 1,
      maxDuration: 365,
      minTravelers: 1,
      maxTravelers: 100,
    },
    budget: {
      minAmount: 0,
      maxAmount: 1000000,
      defaultCurrency: "USD",
    },
    activities: {
      minDuration: 0.25, // 15 minutes
      maxDuration: 24, // hours
      maxPerDay: 10,
    },
  },

  // Reference data for dropdowns
  referenceData: {
    tripTypes: [
      "Leisure",
      "Business",
      "Family",
      "Adventure",
      "Luxury",
      "Budget",
      "Road Trip",
      "Backpacking",
      "Group Tour",
    ],
    accommodationTypes: [
      "Hotel",
      "Airbnb",
      "Resort",
      "Hostel",
      "Vacation Rental",
      "Camping",
    ],
    transportationModes: [
      "Flight",
      "Train",
      "Bus",
      "Rental Car",
      "Taxi/Ride Share",
      "Public Transit",
    ],
    activityCategories: [
      "Sightseeing",
      "Adventure",
      "Cultural",
      "Entertainment",
      "Food & Dining",
      "Shopping",
      "Relaxation",
      "Nature",
    ],
  },

  // Chart configurations
  charts: {
    timeline: {
      width: 800,
      height: 200,
    },
    pie: {
      width: 400,
      height: 300,
    },
    bar: {
      width: 400,
      height: 300,
    },
  },
};


// ================================================================================
// Module: utilities.gs
// ================================================================================

// /src/core/utilities.gs

/**
 * Core utility functions for Travel Budget Planner
 * Dependencies: config.gs
 */
const UtilityModule = {
  /**
   * Applies consistent styling to cell ranges
   * @param {Range} range - Google Sheets range object
   * @param {string} type - Style type ('header', 'input', 'formula', 'warning')
   */
  applyCellStyle(range, type) {
    const styles = CONFIG.styles;

    switch (type) {
      case "header":
        range
          .setBackground(styles.colors.headerBg)
          .setFontColor(styles.colors.headerText)
          .setFontWeight("bold")
          .setFontFamily(styles.fonts.header)
          .setFontSize(12)
          .setHorizontalAlignment("center");
        break;

      case "input":
        range
          .setBackground(styles.colors.inputBg)
          .setFontColor("#000000")
          .setFontFamily(styles.fonts.body)
          .setFontSize(11);
        break;

      case "formula":
        range
          .setBackground(styles.colors.formulaBg)
          .setFontColor("#666666")
          .setFontFamily(styles.fonts.body)
          .setFontSize(11)
          .setNote(
            "This cell contains a formula. Please do not edit directly."
          );
        break;

      case "warning":
        range
          .setBackground(styles.colors.warningBg)
          .setFontColor("#CC0000")
          .setFontFamily(styles.fonts.body)
          .setFontSize(11);
        break;
    }

    // Add borders
    range.setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      styles.colors.borderColor,
      SpreadsheetApp.BorderStyle.SOLID
    );
  },

  /**
   * Creates and applies data validation rules
   * @param {Sheet} sheet - Google Sheet object
   * @param {string} range - Range in A1 notation
   * @param {string} type - Validation type
   * @param {object} options - Additional validation options
   */
  addDataValidation(sheet, range, type, options = {}) {
    let rule;

    switch (type) {
      case "date":
        rule = SpreadsheetApp.newDataValidation()
          .requireDate()
          .setAllowInvalid(false)
          .build();
        break;

      case "number":
        rule = SpreadsheetApp.newDataValidation()
          .requireNumberBetween(
            options.min || 0,
            options.max || Number.MAX_SAFE_INTEGER
          )
          .setAllowInvalid(false)
          .build();
        break;

      case "list":
        rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(options.values, options.strictValidation)
          .setAllowInvalid(false)
          .build();
        break;

      case "checkbox":
        rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        break;

      case "custom":
        rule = SpreadsheetApp.newDataValidation()
          .requireFormulaSatisfied(options.formula)
          .setAllowInvalid(false)
          .build();
        break;
    }

    sheet.getRange(range).setDataValidation(rule);
  },

  /**
   * Formats dates consistently across sheets
   * @param {Date} date - Date to format
   * @returns {string} Formatted date string
   */
  formatDate(date) {
    return Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "MM/dd/yyyy"
    );
  },

  /**
   * Formats currency values
   * @param {number} amount - Amount to format
   * @param {string} currency - Currency code (default: USD)
   * @returns {string} Formatted currency string
   */
  formatCurrency(amount, currency = CONFIG.validation.budget.defaultCurrency) {
    return Utilities.formatString("%s %,.2f", currency, amount);
  },

  /**
   * Validates date ranges
   * @param {Date} startDate - Start date
   * @param {Date} endDate - End date
   * @returns {boolean} True if valid range
   */
  validateDateRange(startDate, endDate) {
    const daysDiff = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24));
    return (
      daysDiff >= CONFIG.validation.trip.minDuration &&
      daysDiff <= CONFIG.validation.trip.maxDuration
    );
  },

  /**
   * Creates named ranges for reference data
   * @param {Spreadsheet} ss - Spreadsheet object
   * @param {Sheet} sheet - Sheet containing reference data
   * @param {string} rangeName - Name for the range
   * @param {string} range - Range in A1 notation
   */
  createNamedRange(ss, sheet, rangeName, range) {
    const existingRange = ss.getRangeByName(rangeName);
    if (existingRange) {
      ss.removeNamedRange(rangeName);
    }
    ss.setNamedRange(rangeName, sheet.getRange(range));
  },
};


// ================================================================================
// Module: protection.gs
// ================================================================================

// /src/core/protection.gs

/**
 * Protection utilities for Travel Budget Planner
 * Dependencies: config.gs
 */
const ProtectionModule = {
  /**
   * Protects entire sheets from editing
   * @param {Spreadsheet} ss - Spreadsheet object
   * @param {string[]} sheetNames - Array of sheet names to protect
   */
  protectSheets(ss, sheetNames) {
    sheetNames.forEach((name) => {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        const protection = sheet.protect();
        protection.setDescription(`Protected sheet: ${name}`);
        protection.setWarningOnly(true);
      }
    });
  },

  /**
   * Protects specific ranges within a sheet
   * @param {Sheet} sheet - Sheet containing ranges to protect
   * @param {string[]} ranges - Array of A1 notation ranges
   * @param {string} description - Protection description
   * @param {boolean} warningOnly - If true, shows warning but allows edits
   */
  protectRanges(sheet, ranges, description, warningOnly = true) {
    ranges.forEach((range) => {
      const protection = sheet.getRange(range).protect();
      protection.setDescription(description);

      if (warningOnly) {
        protection.setWarningOnly(true);
      } else {
        // Only modify editors if not warning only
        const me = Session.getEffectiveUser();
        protection.addEditor(me);
        protection.removeEditors(protection.getEditors());
      }
    });
  },

  /**
   * Protects formula cells in a range
   * @param {Sheet} sheet - Sheet containing formulas
   * @param {string} range - Range in A1 notation
   */
  protectFormulas(sheet, range) {
    const rangeObj = sheet.getRange(range);
    const formulas = rangeObj.getFormulas();

    // Create array of ranges containing formulas
    const formulaRanges = [];
    formulas.forEach((row, i) => {
      row.forEach((cell, j) => {
        if (cell) {
          formulaRanges.push(
            sheet
              .getRange(rangeObj.getRow() + i, rangeObj.getColumn() + j)
              .getA1Notation()
          );
        }
      });
    });

    if (formulaRanges.length > 0) {
      this.protectRanges(
        sheet,
        formulaRanges,
        "Protected formula cells",
        true // Use warning only for formulas
      );
    }
  },

  /**
   * Locks specific columns while allowing row editing
   * @param {Sheet} sheet - Sheet to protect
   * @param {number[]} columns - Array of column numbers to protect
   * @param {string} description - Protection description
   */
  protectColumns(sheet, columns, description) {
    const protection = sheet.protect();
    protection.setDescription(description);
    protection.setWarningOnly(true); // Use warning only for column protection
  },

  /**
   * Adds edit protection while preserving data validation
   * @param {Sheet} sheet - Sheet to protect
   * @param {string} range - Range in A1 notation
   * @param {string} description - Protection description
   */
  protectWithValidation(sheet, range, description) {
    const rangeObj = sheet.getRange(range);
    const validation = rangeObj.getDataValidation();
    const protection = rangeObj.protect();

    protection.setDescription(description);
    protection.setWarningOnly(true); // Use warning only to preserve data validation

    if (validation) {
      rangeObj.setDataValidation(validation);
    }
  },
};


// ================================================================================
// Module: menu.gs
// ================================================================================

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


// ================================================================================
// Module: instructions.gs
// ================================================================================

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


// ================================================================================
// Module: dataTables.gs
// ================================================================================

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


// ================================================================================
// Module: tripSettings.gs
// ================================================================================

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


// ================================================================================
// Module: lodging.gs
// ================================================================================

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


// ================================================================================
// Module: transportation.gs
// ================================================================================

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


// ================================================================================
// Module: activities.gs
// ================================================================================

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


// ================================================================================
// Module: dashboard.gs
// ================================================================================

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


// ================================================================================
// Module: main.gs
// ================================================================================

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

