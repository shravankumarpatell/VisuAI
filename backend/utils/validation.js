// const path = require('path');

// class ValidationUtils {
//   // Validate file extension
//   static validateFileExtension(filename, allowedExtensions) {
//     const ext = path.extname(filename).toLowerCase();
//     return allowedExtensions.includes(ext);
//   }

//   // Validate file size (in MB)
//   static validateFileSize(fileSize, maxSizeMB) {
//     const maxSizeBytes = maxSizeMB * 1024 * 1024;
//     return fileSize <= maxSizeBytes;
//   }

//   // Validate Excel sheet structure
//   static validateExcelStructure(workbook) {
//     const errors = [];
    
//     // Check if workbook has sheets
//     if (!workbook.worksheets || workbook.worksheets.length === 0) {
//       errors.push('Excel file has no worksheets');
//       return { valid: false, errors };
//     }

//     // Check for required sheets
//     const sheetNames = workbook.worksheets.map(ws => ws.name.toLowerCase());
//     const hasTrialGroup = sheetNames.some(name => 
//       name.includes('trial') && name.includes('group')
//     );
//     const hasControlGroup = sheetNames.some(name => 
//       name.includes('control') && name.includes('group')
//     );

//     if (!hasTrialGroup) {
//       errors.push('Missing "Trial Group" sheet');
//     }
//     if (!hasControlGroup) {
//       errors.push('Missing "Control Group" sheet');
//     }

//     // Check if sheets have data
//     workbook.worksheets.forEach(worksheet => {
//       if (worksheet.rowCount < 2) {
//         errors.push(`Sheet "${worksheet.name}" has no data rows`);
//       }
//       if (worksheet.columnCount < 3) {
//         errors.push(`Sheet "${worksheet.name}" has insufficient columns (minimum 3 required)`);
//       }
//     });

//     return {
//       valid: errors.length === 0,
//       errors
//     };
//   }

//   // COMPLETED METHOD: Validate Signs & Symptoms document
//   static validateSignsDocument(tables) {
//     const errors = [];

//     if (!tables || tables.length === 0) {
//       errors.push('No Signs & Symptoms tables found in the Word document');
//       return { valid: false, errors };
//     }

//     tables.forEach((table, index) => {
//       if (!table.title) {
//         errors.push(`Table ${index + 1} is missing a title`);
//       }

//       if (!table.signs || !Array.isArray(table.signs) || table.signs.length === 0) {
//         errors.push(`Table ${index + 1} has no signs data`);
//       }

//       if (!table.assessments || !Array.isArray(table.assessments) || table.assessments.length === 0) {
//         errors.push(`Table ${index + 1} has no assessment periods`);
//       }

//       if (!table.data || typeof table.data !== 'object') {
//         errors.push(`Table ${index + 1} has no percentage data`);
//       } else {
//         // Validate data structure
//         const expectedSigns = ['vedana', 'varna', 'sraava', 'gandha', 'maamasankura', 'parimana'];
//         const expectedAssessments = ['7th day', '14th day', '21st day', '28th day'];

//         table.signs.forEach(sign => {
//           const normalizedSign = sign.toLowerCase();
//           if (!table.data[sign]) {
//             errors.push(`Table ${index + 1}: Missing data for sign "${sign}"`);
//           } else {
//             // Check if all assessment periods have data
//             table.assessments.forEach(assessment => {
//               const value = table.data[sign][assessment];
//               if (typeof value !== 'number' || isNaN(value) || value < 0 || value > 100) {
//                 errors.push(`Table ${index + 1}: Invalid percentage value for ${sign} at ${assessment}: ${value}`);
//               }
//             });
//           }
//         });

//         // Validate that we have reasonable assessment periods
//         if (table.assessments.length < 2) {
//           errors.push(`Table ${index + 1}: Insufficient assessment periods (minimum 2 required)`);
//         }

//         // Check for common assessment period patterns
//         const hasValidAssessmentPattern = table.assessments.some(period => {
//           const normalizedPeriod = period.toLowerCase().replace(/\s+/g, ' ').trim();
//           return normalizedPeriod.includes('day') || normalizedPeriod.includes('week');
//         });

//         if (!hasValidAssessmentPattern) {
//           errors.push(`Table ${index + 1}: Assessment periods should contain time references (e.g., "7 day", "14 day")`);
//         }
//       }

//       // Validate table type
//       if (!table.type || table.type !== 'signs_symptoms') {
//         errors.push(`Table ${index + 1}: Invalid table type, expected "signs_symptoms"`);
//       }

//       // Validate table number
//       if (typeof table.tableNumber !== 'number' || table.tableNumber < 1) {
//         errors.push(`Table ${index + 1}: Invalid or missing table number`);
//       }
//     });

//     return {
//       valid: errors.length === 0,
//       errors
//     };
//   }

//   // Validate Word document tables (existing method)
//   static validateWordTables(tables) {
//     const errors = [];

//     if (!tables || tables.length === 0) {
//       errors.push('No tables found in the Word document');
//       return { valid: false, errors };
//     }

//     tables.forEach((table, index) => {
//       if (!table.title) {
//         errors.push(`Table ${index + 1} is missing a title`);
//       }

//       if (!table.rows || !Array.isArray(table.rows) || table.rows.length === 0) {
//         errors.push(`Table ${index + 1} has no data rows`);
//       }

//       // Validate each row has required fields
//       if (table.rows) {
//         table.rows.forEach((row, rowIndex) => {
//           if (!row.category) {
//             errors.push(`Table ${index + 1}, Row ${rowIndex + 1}: Missing category`);
//           }

//           // Check if numeric values are valid
//           const numericFields = ['trialGroup', 'controlGroup', 'total'];
//           numericFields.forEach(field => {
//             if (row[field] !== undefined && (isNaN(Number(row[field])) || Number(row[field]) < 0)) {
//               errors.push(`Table ${index + 1}, Row ${rowIndex + 1}: Invalid value for ${field}: ${row[field]}`);
//             }
//           });
//         });

//         // Check for at least one data row (excluding total row)
//         const dataRows = table.rows.filter(row => 
//           row.category && row.category.toLowerCase() !== 'total'
//         );
        
//         if (dataRows.length === 0) {
//           errors.push(`Table ${index + 1}: No valid data rows found`);
//         }
//       }
//     });

//     return {
//       valid: errors.length === 0,
//       errors
//     };
//   }

//   // Utility method to sanitize filenames
//   static sanitizeFilename(filename) {
//     return filename.replace(/[^a-zA-Z0-9.-]/g, '_');
//   }
// }

// module.exports = ValidationUtils;




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