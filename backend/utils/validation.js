const path = require('path');

class ValidationUtils {
  static validateFileExtension(filename, allowedExtensions) {
    const ext = path.extname(filename).toLowerCase();
    return allowedExtensions.includes(ext);
  }

  static validateFileSize(sizeInBytes, maxSizeMB) {
    const maxSizeBytes = maxSizeMB * 1024 * 1024;
    return sizeInBytes <= maxSizeBytes;
  }

  static sanitizeFilename(filename) {
    // Remove dangerous characters and normalize
    return filename
      .replace(/[^a-zA-Z0-9.-]/g, '_')
      .replace(/_{2,}/g, '_')
      .replace(/^_+|_+$/g, '')
      .substring(0, 255); // Limit length
  }

  // Updated validation for flexible Word tables (both paired and single-sheet scenarios)
  static validateWordTables(tables) {
    const errors = [];

    if (!Array.isArray(tables)) {
      errors.push('Tables data must be an array');
      return { valid: false, errors };
    }

    if (tables.length === 0) {
      errors.push('No tables found in the document');
      return { valid: false, errors };
    }

    tables.forEach((table, index) => {
      const tableNum = index + 1;

      // Check basic table structure
      if (!table || typeof table !== 'object') {
        errors.push(`Table ${tableNum}: Invalid table structure`);
        return;
      }

      // Check title
      if (!table.title || typeof table.title !== 'string') {
        errors.push(`Table ${tableNum}: Missing or invalid title`);
      }

      // Check rows array
      if (!Array.isArray(table.rows)) {
        errors.push(`Table ${tableNum}: Missing or invalid rows array`);
        return;
      }

      if (table.rows.length === 0) {
        errors.push(`Table ${tableNum}: No data rows found`);
        return;
      }

      // Validate each row structure - now more flexible for single-sheet scenarios
      table.rows.forEach((row, rowIndex) => {
        if (!row || typeof row !== 'object') {
          errors.push(`Table ${tableNum}, Row ${rowIndex + 1}: Invalid row structure`);
          return;
        }

        // Check required fields
        const requiredFields = ['category', 'trialGroup', 'controlGroup', 'total', 'percentage'];
        requiredFields.forEach(field => {
          if (row[field] === undefined || row[field] === null) {
            errors.push(`Table ${tableNum}, Row ${rowIndex + 1}: Missing field '${field}'`);
          }
        });

        // Validate numeric fields - allow 0 for controlGroup in single-sheet scenarios
        if (typeof row.trialGroup !== 'number' || row.trialGroup < 0) {
          errors.push(`Table ${tableNum}, Row ${rowIndex + 1}: Invalid trialGroup value`);
        }

        if (typeof row.controlGroup !== 'number' || row.controlGroup < 0) {
          errors.push(`Table ${tableNum}, Row ${rowIndex + 1}: Invalid controlGroup value`);
        }

        if (typeof row.total !== 'number' || row.total < 0) {
          errors.push(`Table ${tableNum}, Row ${rowIndex + 1}: Invalid total value`);
        }

        // For single-sheet tables, controlGroup will be 0, and that's valid
        // Check data consistency
        const expectedTotal = row.trialGroup + row.controlGroup;
        if (Math.abs(row.total - expectedTotal) > 0.01) { // Allow for minor floating point differences
          errors.push(`Table ${tableNum}, Row ${rowIndex + 1}: Total (${row.total}) doesn't match sum of trialGroup (${row.trialGroup}) + controlGroup (${row.controlGroup})`);
        }

        // Validate percentage
        const percentageValue = parseFloat(row.percentage);
        if (isNaN(percentageValue) || percentageValue < 0 || percentageValue > 100) {
          errors.push(`Table ${tableNum}, Row ${rowIndex + 1}: Invalid percentage value (${row.percentage})`);
        }
      });

      // Check if there's a total row (should be the last row)
      const lastRow = table.rows[table.rows.length - 1];
      if (!lastRow || lastRow.category.toLowerCase() !== 'total') {
        errors.push(`Table ${tableNum}: Missing total row (should be the last row)`);
      } else {
        // Validate total row calculations
        const dataRows = table.rows.slice(0, -1); // All rows except the total row
        const expectedTrialTotal = dataRows.reduce((sum, row) => sum + row.trialGroup, 0);
        const expectedControlTotal = dataRows.reduce((sum, row) => sum + row.controlGroup, 0);
        const expectedGrandTotal = expectedTrialTotal + expectedControlTotal;

        if (Math.abs(lastRow.trialGroup - expectedTrialTotal) > 0.01) {
          errors.push(`Table ${tableNum}: Total row trialGroup (${lastRow.trialGroup}) doesn't match sum of data rows (${expectedTrialTotal})`);
        }

        if (Math.abs(lastRow.controlGroup - expectedControlTotal) > 0.01) {
          errors.push(`Table ${tableNum}: Total row controlGroup (${lastRow.controlGroup}) doesn't match sum of data rows (${expectedControlTotal})`);
        }

        if (Math.abs(lastRow.total - expectedGrandTotal) > 0.01) {
          errors.push(`Table ${tableNum}: Total row total (${lastRow.total}) doesn't match calculated total (${expectedGrandTotal})`);
        }

        // For single-sheet tables, it's acceptable to have controlGroup = 0
        if (expectedControlTotal === 0 && expectedTrialTotal > 0) {
          console.log(`Table ${tableNum}: Detected single-sheet table (no control group data)`);
        }
      }
    });

    return {
      valid: errors.length === 0,
      errors: errors
    };
  }

  // Validation for Signs documents
  static validateSignsDocument(tables) {
    const errors = [];

    if (!Array.isArray(tables)) {
      errors.push('Signs data must be an array');
      return { valid: false, errors };
    }

    if (tables.length === 0) {
      errors.push('No Signs & Symptoms tables found in the document');
      return { valid: false, errors };
    }

    tables.forEach((table, index) => {
      const tableNum = index + 1;

      // Check basic table structure
      if (!table || typeof table !== 'object') {
        errors.push(`Signs Table ${tableNum}: Invalid table structure`);
        return;
      }

      // Check required fields for Signs & Symptoms data
      if (!table.signs || !Array.isArray(table.signs) || table.signs.length === 0) {
        errors.push(`Signs Table ${tableNum}: Missing or invalid signs array`);
      }

      if (!table.assessments || !Array.isArray(table.assessments) || table.assessments.length === 0) {
        errors.push(`Signs Table ${tableNum}: Missing or invalid assessments array`);
      }

      if (!table.data || typeof table.data !== 'object') {
        errors.push(`Signs Table ${tableNum}: Missing or invalid data object`);
        return;
      }

      // Validate data completeness
      if (table.signs && table.assessments) {
        table.signs.forEach(sign => {
          if (!table.data[sign]) {
            errors.push(`Signs Table ${tableNum}: Missing data for sign '${sign}'`);
            return;
          }

          table.assessments.forEach(assessment => {
            if (typeof table.data[sign][assessment] !== 'number') {
              errors.push(`Signs Table ${tableNum}: Missing or invalid percentage data for '${sign}' at '${assessment}'`);
            } else {
              const percentage = table.data[sign][assessment];
              if (percentage < 0 || percentage > 100) {
                errors.push(`Signs Table ${tableNum}: Invalid percentage value (${percentage}) for '${sign}' at '${assessment}'`);
              }
            }
          });
        });
      }

      // Check table type
      if (table.type && table.type !== 'signs_symptoms') {
        console.warn(`Signs Table ${tableNum}: Unexpected table type '${table.type}'`);
      }
    });

    return {
      valid: errors.length === 0,
      errors: errors
    };
  }

  // Enhanced validation for Excel data with flexible sheet handling
  static validateExcelData(data) {
    const errors = [];

    if (!data || typeof data !== 'object') {
      errors.push('Invalid Excel data structure');
      return { valid: false, errors };
    }

    if (!Array.isArray(data.tables)) {
      errors.push('Excel data must contain a tables array');
      return { valid: false, errors };
    }

    if (data.tables.length === 0) {
      errors.push('No tables found in Excel data');
      return { valid: false, errors };
    }

    // Use the same validation logic as Word tables since the structure is the same
    const wordValidation = this.validateWordTables(data.tables);
    
    if (!wordValidation.valid) {
      errors.push(...wordValidation.errors.map(error => `Excel ${error}`));
    }

    // Additional checks for Excel-specific data
    data.tables.forEach((table, index) => {
      if (table.sourceInfo) {
        console.log(`Table ${index + 1}: Source - ${table.sourceInfo.trialGroup} & ${table.sourceInfo.controlGroup}`);
        
        // Log information about single-sheet tables
        if (table.sourceInfo.controlGroup === 'N/A (Single Sheet)') {
          console.log(`Table ${index + 1}: Single-sheet table detected from ${table.sourceInfo.trialGroup}`);
        }
      }
    });

    return {
      valid: errors.length === 0,
      errors: errors
    };
  }

  // Helper method to check if file exists and is readable
  static async validateFileAccess(filePath) {
    try {
      const fs = require('fs').promises;
      await fs.access(filePath);
      return { valid: true, error: null };
    } catch (error) {
      return { valid: false, error: `File not accessible: ${error.message}` };
    }
  }

  // Enhanced filename validation
  static validateFilename(filename) {
    const errors = [];

    if (!filename || typeof filename !== 'string') {
      errors.push('Filename must be a non-empty string');
      return { valid: false, errors };
    }

    if (filename.length > 255) {
      errors.push('Filename is too long (max 255 characters)');
    }

    // Check for dangerous characters
    const dangerousChars = /[<>:"/\\|?*\x00-\x1f]/;
    if (dangerousChars.test(filename)) {
      errors.push('Filename contains invalid characters');
    }

    // Check for reserved names (Windows)
    const reservedNames = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'];
    const nameWithoutExt = path.parse(filename).name.toUpperCase();
    if (reservedNames.includes(nameWithoutExt)) {
      errors.push('Filename uses a reserved system name');
    }

    return {
      valid: errors.length === 0,
      errors: errors
    };
  }
}

module.exports = ValidationUtils;