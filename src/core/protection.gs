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
