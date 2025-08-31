const path = require('path');

class ValidationUtils {
  // Validate file extension
  static validateFileExtension(filename, allowedExtensions) {
    const ext = path.extname(filename).toLowerCase();
    return allowedExtensions.includes(ext);
  }

  // Validate file size (in MB)
  static validateFileSize(fileSize, maxSizeMB) {
    const maxSizeBytes = maxSizeMB * 1024 * 1024;
    return fileSize <= maxSizeBytes;
  }

  // Validate Excel sheet structure
  static validateExcelStructure(workbook) {
    const errors = [];
    
    // Check if workbook has sheets
    if (!workbook.worksheets || workbook.worksheets.length === 0) {
      errors.push('Excel file has no worksheets');
      return { valid: false, errors };
    }

    // Check for required sheets
    const sheetNames = workbook.worksheets.map(ws => ws.name.toLowerCase());
    const hasTrialGroup = sheetNames.some(name => 
      name.includes('trial') && name.includes('group')
    );
    const hasControlGroup = sheetNames.some(name => 
      name.includes('control') && name.includes('group')
    );

    if (!hasTrialGroup) {
      errors.push('Missing "Trial Group" sheet');
    }
    if (!hasControlGroup) {
      errors.push('Missing "Control Group" sheet');
    }

    // Check if sheets have data
    workbook.worksheets.forEach(worksheet => {
      if (worksheet.rowCount < 2) {
        errors.push(`Sheet "${worksheet.name}" has no data rows`);
      }
      if (worksheet.columnCount < 3) {
        errors.push(`Sheet "${worksheet.name}" has insufficient columns (minimum 3 required)`);
      }
    });

    return {
      valid: errors.length === 0,
      errors
    };
  }

  // COMPLETED METHOD: Validate Signs & Symptoms document
  static validateSignsDocument(tables) {
    const errors = [];

    if (!tables || tables.length === 0) {
      errors.push('No Signs & Symptoms tables found in the Word document');
      return { valid: false, errors };
    }

    tables.forEach((table, index) => {
      if (!table.title) {
        errors.push(`Table ${index + 1} is missing a title`);
      }

      if (!table.signs || !Array.isArray(table.signs) || table.signs.length === 0) {
        errors.push(`Table ${index + 1} has no signs data`);
      }

      if (!table.assessments || !Array.isArray(table.assessments) || table.assessments.length === 0) {
        errors.push(`Table ${index + 1} has no assessment periods`);
      }

      if (!table.data || typeof table.data !== 'object') {
        errors.push(`Table ${index + 1} has no percentage data`);
      } else {
        // Validate data structure
        const expectedSigns = ['vedana', 'varna', 'sraava', 'gandha', 'maamasankura', 'parimana'];
        const expectedAssessments = ['7th day', '14th day', '21st day', '28th day'];

        table.signs.forEach(sign => {
          const normalizedSign = sign.toLowerCase();
          if (!table.data[sign]) {
            errors.push(`Table ${index + 1}: Missing data for sign "${sign}"`);
          } else {
            // Check if all assessment periods have data
            table.assessments.forEach(assessment => {
              const value = table.data[sign][assessment];
              if (typeof value !== 'number' || isNaN(value) || value < 0 || value > 100) {
                errors.push(`Table ${index + 1}: Invalid percentage value for ${sign} at ${assessment}: ${value}`);
              }
            });
          }
        });

        // Validate that we have reasonable assessment periods
        if (table.assessments.length < 2) {
          errors.push(`Table ${index + 1}: Insufficient assessment periods (minimum 2 required)`);
        }

        // Check for common assessment period patterns
        const hasValidAssessmentPattern = table.assessments.some(period => {
          const normalizedPeriod = period.toLowerCase().replace(/\s+/g, ' ').trim();
          return normalizedPeriod.includes('day') || normalizedPeriod.includes('week');
        });

        if (!hasValidAssessmentPattern) {
          errors.push(`Table ${index + 1}: Assessment periods should contain time references (e.g., "7 day", "14 day")`);
        }
      }

      // Validate table type
      if (!table.type || table.type !== 'signs_symptoms') {
        errors.push(`Table ${index + 1}: Invalid table type, expected "signs_symptoms"`);
      }

      // Validate table number
      if (typeof table.tableNumber !== 'number' || table.tableNumber < 1) {
        errors.push(`Table ${index + 1}: Invalid or missing table number`);
      }
    });

    return {
      valid: errors.length === 0,
      errors
    };
  }

  // Validate Word document tables (existing method)
  static validateWordTables(tables) {
    const errors = [];

    if (!tables || tables.length === 0) {
      errors.push('No tables found in the Word document');
      return { valid: false, errors };
    }

    tables.forEach((table, index) => {
      if (!table.title) {
        errors.push(`Table ${index + 1} is missing a title`);
      }

      if (!table.rows || !Array.isArray(table.rows) || table.rows.length === 0) {
        errors.push(`Table ${index + 1} has no data rows`);
      }

      // Validate each row has required fields
      if (table.rows) {
        table.rows.forEach((row, rowIndex) => {
          if (!row.category) {
            errors.push(`Table ${index + 1}, Row ${rowIndex + 1}: Missing category`);
          }

          // Check if numeric values are valid
          const numericFields = ['trialGroup', 'controlGroup', 'total'];
          numericFields.forEach(field => {
            if (row[field] !== undefined && (isNaN(Number(row[field])) || Number(row[field]) < 0)) {
              errors.push(`Table ${index + 1}, Row ${rowIndex + 1}: Invalid value for ${field}: ${row[field]}`);
            }
          });
        });

        // Check for at least one data row (excluding total row)
        const dataRows = table.rows.filter(row => 
          row.category && row.category.toLowerCase() !== 'total'
        );
        
        if (dataRows.length === 0) {
          errors.push(`Table ${index + 1}: No valid data rows found`);
        }
      }
    });

    return {
      valid: errors.length === 0,
      errors
    };
  }

  // Utility method to sanitize filenames
  static sanitizeFilename(filename) {
    return filename.replace(/[^a-zA-Z0-9.-]/g, '_');
  }
}

module.exports = ValidationUtils;