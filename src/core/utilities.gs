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
