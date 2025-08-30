const ExcelJS = require('exceljs');

class ExcelService {
  async readExcelFile(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const data = {
      trialGroup: null,
      controlGroup: null,
      thirdGroup: null,
      fourthGroup: null,
      tables: []
    };

    // Find all four sheets by position
    const sheet1 = this.findSheetByPosition(workbook, 0);  // First Group
    const sheet2 = this.findSheetByPosition(workbook, 1);  // Second Group
    const sheet3 = this.findSheetByPosition(workbook, 2);  // Third Group
    const sheet4 = this.findSheetByPosition(workbook, 3);  // Fourth Group

    if (!sheet1 || !sheet2 || !sheet3 || !sheet4) {
      const actualSheetNames = workbook.worksheets.map(ws => ws.name).join(', ');
      throw new Error(
        `Expected 4 sheets but found ${workbook.worksheets.length}. ` +
        `Found sheets: ${actualSheetNames}. ` +
        `Please ensure you have exactly 4 sheets in your Excel file.`
      );
    }

    console.log(`Using sheets: "${sheet1.name}", "${sheet2.name}", "${sheet3.name}", "${sheet4.name}"`);

    // Read data from all four sheets
    const sheet1Data = this.extractSheetData(sheet1);
    const sheet2Data = this.extractSheetData(sheet2);
    const sheet3Data = this.extractSheetData(sheet3);
    const sheet4Data = this.extractSheetData(sheet4);

    // Validate that related pairs have same number of columns
    if (sheet1Data.headers.length !== sheet2Data.headers.length) {
      throw new Error('Sheet 1 and Sheet 2 must have the same number of columns');
    }

    if (sheet3Data.headers.length !== sheet4Data.headers.length) {
      throw new Error('Sheet 3 and Sheet 4 must have the same number of columns');
    }

    // Generate tables for Sheet 1 & Sheet 2 pair (excluding first 2 columns)
    for (let colIndex = 2; colIndex < sheet1Data.headers.length; colIndex++) {
      const header = sheet1Data.headers[colIndex];
      
      const sheet1Values = this.getColumnValues(sheet1Data.rows, colIndex);
      const sheet2Values = this.getColumnValues(sheet2Data.rows, colIndex);
      
      const table = this.createTable(header, sheet1Values, sheet2Values);
      data.tables.push(table);
    }

    // Generate tables for Sheet 3 & Sheet 4 pair (excluding first 2 columns)
    for (let colIndex = 2; colIndex < sheet3Data.headers.length; colIndex++) {
      const header = sheet3Data.headers[colIndex];
      
      const sheet3Values = this.getColumnValues(sheet3Data.rows, colIndex);
      const sheet4Values = this.getColumnValues(sheet4Data.rows, colIndex);
      
      const table = this.createTable(header, sheet3Values, sheet4Values);
      data.tables.push(table);
    }

    return data;
  }

  findSheet(workbook, sheetNames) {
    for (const worksheet of workbook.worksheets) {
      const sheetName = worksheet.name.toLowerCase();
      if (sheetNames.some(name => sheetName.includes(name))) {
        return worksheet;
      }
    }
    return null;
  }

  // Alternative method to find sheets by position if names don't match
  findSheetByPosition(workbook, position) {
    if (workbook.worksheets && workbook.worksheets.length > position) {
      return workbook.worksheets[position];
    }
    return null;
  }

  extractSheetData(worksheet) {
    const data = {
      headers: [],
      rows: []
    };

    let headerRow = null;
    let dataStartRow = null;

    // Find header row (usually the first non-empty row)
    worksheet.eachRow((row, rowNumber) => {
      if (!headerRow && row.actualCellCount > 0) {
        headerRow = rowNumber;
        row.eachCell((cell, colNumber) => {
          data.headers[colNumber - 1] = cell.value ? cell.value.toString() : '';
        });
      } else if (headerRow && !dataStartRow) {
        dataStartRow = rowNumber;
      }
    });

    // Extract data rows
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber >= dataStartRow) {
        const rowData = [];
        row.eachCell((cell, colNumber) => {
          rowData[colNumber - 1] = cell.value;
        });
        if (rowData.some(val => val !== null && val !== undefined && val !== '')) {
          data.rows.push(rowData);
        }
      }
    });

    return data;
  }

  getColumnValues(rows, colIndex) {
    const values = {};
    
    rows.forEach(row => {
      const value = row[colIndex];
      if (value !== null && value !== undefined && value !== '') {
        // Normalize the key to handle case variations and extra spaces
        const key = value.toString().trim().toLowerCase();
        const originalValue = value.toString().trim();
        
        // Use the original casing for display, but normalize for counting
        if (values[key]) {
          values[key].count += 1;
        } else {
          values[key] = {
            count: 1,
            displayValue: originalValue
          };
        }
      }
    });

    // Convert back to the expected format
    const result = {};
    Object.keys(values).forEach(key => {
      result[values[key].displayValue] = values[key].count;
    });

    return result;
  }

  createTable(header, trialValues, controlValues) {
    // Get all unique values (normalized to avoid duplicates)
    const allValuesSet = new Set();
    const valueMapping = {};
    
    // Process trial values
    Object.keys(trialValues).forEach(value => {
      const normalizedKey = value.trim().toLowerCase();
      allValuesSet.add(normalizedKey);
      if (!valueMapping[normalizedKey]) {
        valueMapping[normalizedKey] = {
          displayValue: value.trim(),
          trialCount: 0,
          controlCount: 0
        };
      }
      valueMapping[normalizedKey].trialCount += trialValues[value];
    });
    
    // Process control values
    Object.keys(controlValues).forEach(value => {
      const normalizedKey = value.trim().toLowerCase();
      allValuesSet.add(normalizedKey);
      if (!valueMapping[normalizedKey]) {
        valueMapping[normalizedKey] = {
          displayValue: value.trim(),
          trialCount: 0,
          controlCount: 0
        };
      }
      valueMapping[normalizedKey].controlCount += controlValues[value];
    });

    const rows = [];
    let trialTotal = 0;
    let controlTotal = 0;

    // Create rows for each unique value
    Array.from(allValuesSet).sort().forEach(normalizedKey => {
      const valueData = valueMapping[normalizedKey];
      const trialCount = valueData.trialCount;
      const controlCount = valueData.controlCount;
      const total = trialCount + controlCount;
      
      trialTotal += trialCount;
      controlTotal += controlCount;

      rows.push({
        category: valueData.displayValue,
        trialGroup: trialCount,
        controlGroup: controlCount,
        total: total,
        percentage: 0 // Will be calculated after we have the grand total
      });
    });

    const grandTotal = trialTotal + controlTotal;

    // Calculate percentages
    rows.forEach(row => {
      row.percentage = grandTotal > 0 ? ((row.total / grandTotal) * 100).toFixed(2) : '0.00';
    });

    // Add total row
    rows.push({
      category: 'Total',
      trialGroup: trialTotal,
      controlGroup: controlTotal,
      total: grandTotal,
      percentage: '100.00'
    });

    return {
      title: header,
      rows: rows
    };
  }
}

module.exports = new ExcelService();