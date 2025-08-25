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

  // Validate Word document tables
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

      if (!table.rows || table.rows.length === 0) {
        errors.push(`Table ${index + 1} has no data rows`);
      } else {
        // Check table structure
        const requiredFields = ['category', 'trialGroup', 'controlGroup', 'total', 'percentage'];
        const sampleRow = table.rows[0];
        
        requiredFields.forEach(field => {
          if (!(field in sampleRow)) {
            errors.push(`Table ${index + 1} is missing required field: ${field}`);
          }
        });

        // Validate data types
        table.rows.forEach((row, rowIndex) => {
          if (typeof row.trialGroup !== 'number' || isNaN(row.trialGroup)) {
            errors.push(`Table ${index + 1}, Row ${rowIndex + 1}: Invalid trial group value`);
          }
          if (typeof row.controlGroup !== 'number' || isNaN(row.controlGroup)) {
            errors.push(`Table ${index + 1}, Row ${rowIndex + 1}: Invalid control group value`);
          }
        });
      }
    });

    return {
      valid: errors.length === 0,
      errors
    };
  }

  // Sanitize filename
  static sanitizeFilename(filename) {
    // Remove path traversal attempts
    let sanitized = path.basename(filename);
    
    // Remove special characters except dots and hyphens
    sanitized = sanitized.replace(/[^a-zA-Z0-9.-]/g, '_');
    
    // Ensure it doesn't start with a dot
    if (sanitized.startsWith('.')) {
      sanitized = '_' + sanitized.substring(1);
    }
    
    return sanitized;
  }

  // Validate table data consistency
  static validateTableDataConsistency(tableData) {
    const errors = [];
    
    tableData.rows.forEach((row, index) => {
      // Skip total row
      if (row.category.toLowerCase() === 'total') {
        return;
      }
      
      // Check if total equals sum of trial and control
      const calculatedTotal = row.trialGroup + row.controlGroup;
      if (calculatedTotal !== row.total) {
        errors.push(`Row ${index + 1}: Total (${row.total}) doesn't match sum of Trial (${row.trialGroup}) + Control (${row.controlGroup})`);
      }
    });
    
    // Validate percentages sum to 100
    const totalRow = tableData.rows.find(row => row.category.toLowerCase() === 'total');
    if (totalRow && Math.abs(parseFloat(totalRow.percentage) - 100) > 0.01) {
      errors.push('Total percentage does not equal 100%');
    }
    
    return {
      valid: errors.length === 0,
      errors
    };
  }
}

module.exports = ValidationUtils;