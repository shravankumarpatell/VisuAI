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

  // NEW: Validation for Improvement Document data
  static validateImprovementDocument(improvementData) {
    const errors = [];

    if (!improvementData || typeof improvementData !== 'object') {
      errors.push('Invalid improvement document data structure');
      return { valid: false, errors };
    }

    // Check basic structure
    if (!Array.isArray(improvementData.categories)) {
      errors.push('Improvement data must contain a categories array');
      return { valid: false, errors };
    }

    if (!Array.isArray(improvementData.timePeriods)) {
      errors.push('Improvement data must contain a timePeriods array');
      return { valid: false, errors };
    }

    if (!improvementData.groupA || typeof improvementData.groupA !== 'object') {
      errors.push('Improvement data must contain Group A data');
      return { valid: false, errors };
    }

    if (!improvementData.groupB || typeof improvementData.groupB !== 'object') {
      errors.push('Improvement data must contain Group B data');
      return { valid: false, errors };
    }

    // Validate categories array
    if (improvementData.categories.length === 0) {
      errors.push('No improvement categories found');
      return { valid: false, errors };
    }

    // Expected improvement categories
    const expectedCategories = [
      'Cured(100%)',
      'Marked improved(75-100%)',
      'Moderate improved(50-75%)',
      'Mild improved(25-50%)',
      'Not cured(<25%)'
    ];

    // Check if categories are reasonable (allow some flexibility)
    const hasValidCategories = improvementData.categories.some(cat => 
      expectedCategories.some(expected => 
        cat.toLowerCase().includes(expected.toLowerCase().split('(')[0]) ||
        expected.toLowerCase().includes(cat.toLowerCase().split('(')[0])
      )
    );

    if (!hasValidCategories) {
      errors.push('No recognizable improvement categories found. Expected categories like: Cured, Marked improved, Moderate improved, Mild improved, Not cured');
    }

    // Validate time periods array
    if (improvementData.timePeriods.length === 0) {
      errors.push('No time periods found');
      return { valid: false, errors };
    }

    // Expected time periods (allow flexibility)
    const expectedTimePeriods = ['7th', '14th', '21st', '28th'];
    const hasValidTimePeriods = improvementData.timePeriods.some(period => 
      expectedTimePeriods.some(expected => 
        period.toString().includes(expected.replace('th', '')) ||
        period.toString().includes(expected)
      )
    );

    if (!hasValidTimePeriods) {
      errors.push('No recognizable time periods found. Expected periods like: 7th, 14th, 21st, 28th day');
    }

    // Validate Group A data structure
    improvementData.timePeriods.forEach(period => {
      if (!improvementData.groupA[period]) {
        errors.push(`Missing Group A data for time period: ${period}`);
        return;
      }

      improvementData.categories.forEach(category => {
        const categoryData = improvementData.groupA[period][category];
        
        if (!categoryData) {
          errors.push(`Missing Group A data for category '${category}' at time period '${period}'`);
        } else {
          // Validate count and percentage structure
          if (typeof categoryData.count !== 'number' || categoryData.count < 0) {
            errors.push(`Invalid count value for Group A '${category}' at '${period}': ${categoryData.count}`);
          }

          if (typeof categoryData.percentage !== 'number' || categoryData.percentage < 0 || categoryData.percentage > 100) {
            errors.push(`Invalid percentage value for Group A '${category}' at '${period}': ${categoryData.percentage}`);
          }
        }
      });
    });

    // Validate Group B data structure
    improvementData.timePeriods.forEach(period => {
      if (!improvementData.groupB[period]) {
        errors.push(`Missing Group B data for time period: ${period}`);
        return;
      }

      improvementData.categories.forEach(category => {
        const categoryData = improvementData.groupB[period][category];
        
        if (!categoryData) {
          errors.push(`Missing Group B data for category '${category}' at time period '${period}'`);
        } else {
          // Validate count and percentage structure
          if (typeof categoryData.count !== 'number' || categoryData.count < 0) {
            errors.push(`Invalid count value for Group B '${category}' at '${period}': ${categoryData.count}`);
          }

          if (typeof categoryData.percentage !== 'number' || categoryData.percentage < 0 || categoryData.percentage > 100) {
            errors.push(`Invalid percentage value for Group B '${category}' at '${period}': ${categoryData.percentage}`);
          }
        }
      });
    });

    // Validate data consistency within each group and time period
    improvementData.timePeriods.forEach(period => {
      // Check Group A totals
      let groupATotal = 0;
      improvementData.categories.forEach(category => {
        if (improvementData.groupA[period] && improvementData.groupA[period][category]) {
          groupATotal += improvementData.groupA[period][category].count;
        }
      });

      // Check Group B totals
      let groupBTotal = 0;
      improvementData.categories.forEach(category => {
        if (improvementData.groupB[period] && improvementData.groupB[period][category]) {
          groupBTotal += improvementData.groupB[period][category].count;
        }
      });

      // Warn if totals seem inconsistent (but don't fail validation)
      if (groupATotal === 0 && groupBTotal === 0) {
        console.warn(`Warning: No data found for time period '${period}'`);
      }

      // Check for reasonable sample sizes (warn, don't fail)
      if (groupATotal > 100 || groupBTotal > 100) {
        console.warn(`Warning: Large sample size detected at '${period}' (Group A: ${groupATotal}, Group B: ${groupBTotal})`);
      }
    });

    // Validate minimum data requirements
    let hasAnyData = false;
    improvementData.timePeriods.forEach(period => {
      improvementData.categories.forEach(category => {
        const groupAData = improvementData.groupA[period] && improvementData.groupA[period][category];
        const groupBData = improvementData.groupB[period] && improvementData.groupB[period][category];
        
        if ((groupAData && groupAData.count > 0) || (groupBData && groupBData.count > 0)) {
          hasAnyData = true;
        }
      });
    });

    if (!hasAnyData) {
      errors.push('No improvement data found. All counts appear to be zero or missing.');
    }

    return {
      valid: errors.length === 0,
      errors: errors
    };
  }

  // NEW: Validation for Master Chart data
  static validateMasterChartData(masterChartData) {
    const errors = [];

    if (!masterChartData || typeof masterChartData !== 'object') {
      errors.push('Invalid master chart data structure');
      return { valid: false, errors };
    }

    // Check basic structure
    if (!Array.isArray(masterChartData.tables)) {
      errors.push('Master chart data must contain a tables array');
      return { valid: false, errors };
    }

    if (masterChartData.tables.length === 0) {
      errors.push('No tables found in master chart data');
      return { valid: false, errors };
    }

    // Check metadata
    if (!masterChartData.sourceFile) {
      errors.push('Missing source file information');
    }

    if (!masterChartData.totalSheets || masterChartData.totalSheets < 1) {
      errors.push('Invalid total sheets count');
    }

    if (!masterChartData.processedSheets || masterChartData.processedSheets < 1) {
      errors.push('No sheets were successfully processed');
    }

    // Validate each table
    masterChartData.tables.forEach((table, index) => {
      const tableNum = index + 1;

      // Check basic table structure
      if (!table || typeof table !== 'object') {
        errors.push(`Table ${tableNum}: Invalid table structure`);
        return;
      }

      // Check required fields
      if (!table.title || typeof table.title !== 'string') {
        errors.push(`Table ${tableNum}: Missing or invalid title`);
      }

      if (typeof table.sheetNumber !== 'number' || table.sheetNumber < 1) {
        errors.push(`Table ${tableNum}: Invalid sheet number`);
      }

      if (!table.sheetName || typeof table.sheetName !== 'string') {
        errors.push(`Table ${tableNum}: Missing sheet name`);
      }

      // Check parameter group structure
      if (!table.parameterGroup || typeof table.parameterGroup !== 'object') {
        errors.push(`Table ${tableNum}: Missing parameter group information`);
        return;
      }

      const paramGroup = table.parameterGroup;
      if (!paramGroup.name || typeof paramGroup.name !== 'string') {
        errors.push(`Table ${tableNum}: Missing parameter name`);
      }

      if (typeof paramGroup.btIndex !== 'number' || paramGroup.btIndex < 0) {
        errors.push(`Table ${tableNum}: Invalid BT column index`);
      }

      if (typeof paramGroup.atIndex !== 'number' || paramGroup.atIndex < 0) {
        errors.push(`Table ${tableNum}: Invalid AT column index`);
      }

      if (!paramGroup.btHeader || !paramGroup.atHeader) {
        errors.push(`Table ${tableNum}: Missing BT/AT column headers`);
      }

      // Check raw data structure
      if (!table.rawData || typeof table.rawData !== 'object') {
        errors.push(`Table ${tableNum}: Missing raw data`);
        return;
      }

      const rawData = table.rawData;
      if (!Array.isArray(rawData.bt) || !Array.isArray(rawData.at)) {
        errors.push(`Table ${tableNum}: Invalid BT/AT data arrays`);
        return;
      }

      if (rawData.bt.length !== rawData.at.length) {
        errors.push(`Table ${tableNum}: BT and AT arrays must have the same length`);
      }

      if (rawData.bt.length === 0) {
        errors.push(`Table ${tableNum}: No valid data points found`);
      }

      if (typeof rawData.n !== 'number' || rawData.n !== rawData.bt.length) {
        errors.push(`Table ${tableNum}: Invalid sample size (n)`);
      }

      // Check minimum sample size for statistical validity
      if (rawData.n < 3) {
        errors.push(`Table ${tableNum}: Sample size too small for statistical analysis (minimum 3 required, got ${rawData.n})`);
      }

      // Validate that BT/AT data contains only numeric values
      rawData.bt.forEach((value, i) => {
        if (typeof value !== 'number' || !isFinite(value)) {
          errors.push(`Table ${tableNum}: Invalid BT value at position ${i + 1}: ${value}`);
        }
      });

      rawData.at.forEach((value, i) => {
        if (typeof value !== 'number' || !isFinite(value)) {
          errors.push(`Table ${tableNum}: Invalid AT value at position ${i + 1}: ${value}`);
        }
      });

      // Check statistics structure
      if (!table.statistics || typeof table.statistics !== 'object') {
        errors.push(`Table ${tableNum}: Missing statistics`);
        return;
      }

      const stats = table.statistics;
      
      // Check required statistical fields
      const requiredStats = [
        'n', 'meanBT', 'meanAT', 'meanDiff', 'sdBT', 'sdAT', 'sdDiff',
        'standardError', 'tValue', 'degreesOfFreedom', 'pValue', 
        'effectiveness', 'significance'
      ];

      requiredStats.forEach(field => {
        if (!(field in stats)) {
          errors.push(`Table ${tableNum}: Missing statistical field '${field}'`);
        } else if (field !== 'significance' && (typeof stats[field] !== 'number' || !isFinite(stats[field]))) {
          errors.push(`Table ${tableNum}: Invalid ${field} value: ${stats[field]}`);
        }
      });

      // Validate significance structure
      if (stats.significance) {
        if (typeof stats.significance !== 'object') {
          errors.push(`Table ${tableNum}: Invalid significance structure`);
        } else {
          const sig = stats.significance;
          if (!sig.level || !sig.description || !sig.symbol) {
            errors.push(`Table ${tableNum}: Incomplete significance information`);
          }

          // Check valid significance levels
          const validLevels = ['HS', 'S', 'VS', 'NS'];
          if (!validLevels.includes(sig.level)) {
            errors.push(`Table ${tableNum}: Invalid significance level '${sig.level}'`);
          }
        }
      }

      // Validate statistical consistency
      if (stats.n && rawData.n && stats.n !== rawData.n) {
        errors.push(`Table ${tableNum}: Inconsistent sample sizes (stats.n=${stats.n}, rawData.n=${rawData.n})`);
      }

      if (stats.degreesOfFreedom && stats.n && stats.degreesOfFreedom !== stats.n - 1) {
        errors.push(`Table ${tableNum}: Invalid degrees of freedom (should be n-1 = ${stats.n - 1}, got ${stats.degreesOfFreedom})`);
      }

      // Check p-value range
      if (stats.pValue && (stats.pValue < 0 || stats.pValue > 1)) {
        errors.push(`Table ${tableNum}: Invalid p-value range (${stats.pValue}), should be between 0 and 1`);
      }

      // Check raw differences array if present
      if (stats.rawDifferences) {
        if (!Array.isArray(stats.rawDifferences)) {
          errors.push(`Table ${tableNum}: Raw differences must be an array`);
        } else if (stats.rawDifferences.length !== rawData.n) {
          errors.push(`Table ${tableNum}: Raw differences array length doesn't match sample size`);
        }
      }
    });

    // Check for duplicate parameter names within same sheet
    const sheetParameterMap = new Map();
    masterChartData.tables.forEach((table, index) => {
      const key = `${table.sheetNumber}-${table.parameterGroup.name}`;
      if (sheetParameterMap.has(key)) {
        errors.push(`Duplicate parameter '${table.parameterGroup.name}' found in sheet ${table.sheetNumber}`);
      }
      sheetParameterMap.set(key, true);
    });

    return {
      valid: errors.length === 0,
      errors: errors
    };
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