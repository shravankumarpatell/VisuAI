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

    console.log(`Found ${workbook.worksheets.length} sheets in the Excel file`);

    // Find all available sheets by position (up to 4 sheets)
    const sheet1 = this.findSheetByPosition(workbook, 0);  // First Group (Trial)
    const sheet2 = this.findSheetByPosition(workbook, 1);  // Second Group (Control)
    const sheet3 = this.findSheetByPosition(workbook, 2);  // Third Group
    const sheet4 = this.findSheetByPosition(workbook, 3);  // Fourth Group

    const availableSheets = [];
    if (sheet1) availableSheets.push({ sheet: sheet1, name: 'Sheet 1', position: 1 });
    if (sheet2) availableSheets.push({ sheet: sheet2, name: 'Sheet 2', position: 2 });
    if (sheet3) availableSheets.push({ sheet: sheet3, name: 'Sheet 3', position: 3 });
    if (sheet4) availableSheets.push({ sheet: sheet4, name: 'Sheet 4', position: 4 });

    if (availableSheets.length === 0) {
      throw new Error('No sheets found in the Excel file.');
    }

    console.log(`Available sheets: ${availableSheets.map(s => s.name + ' (' + s.sheet.name + ')').join(', ')}`);

    // Extract data from available sheets and check if they have meaningful data
    const sheetDataArray = [];
    for (const sheetInfo of availableSheets) {
      try {
        const sheetData = this.extractSheetData(sheetInfo.sheet);
        const hasData = this.hasValidData(sheetData);
        
        sheetDataArray.push({
          ...sheetInfo,
          data: sheetData,
          hasData: hasData
        });
        
        console.log(`${sheetInfo.name}: ${hasData ? 'HAS DATA' : 'EMPTY/NO DATA'} - ${sheetData.rows.length} rows`);
      } catch (error) {
        console.warn(`Error processing ${sheetInfo.name}:`, error.message);
        sheetDataArray.push({
          ...sheetInfo,
          data: null,
          hasData: false
        });
      }
    }

    // Filter sheets with valid data
    const validSheets = sheetDataArray.filter(sheet => sheet.hasData);
    
    if (validSheets.length === 0) {
      throw new Error('No sheets contain valid data. Please ensure your Excel file has data in at least one sheet.');
    }

    console.log(`Valid sheets with data: ${validSheets.map(s => s.name).join(', ')}`);

    // Generate tables based on available data
    // Strategy: Try to pair sheets, but if pairing fails, create single-sheet tables
    
    // First, try traditional pairing (1+2, 3+4)
    const sheet1Data = sheetDataArray.find(s => s.position === 1);
    const sheet2Data = sheetDataArray.find(s => s.position === 2);
    const sheet3Data = sheetDataArray.find(s => s.position === 3);
    const sheet4Data = sheetDataArray.find(s => s.position === 4);

    // Process Sheet 1 & 2 pair
    if (sheet1Data && sheet1Data.hasData) {
      if (sheet2Data && sheet2Data.hasData) {
        // Both sheets available - create paired tables
        console.log('Creating paired tables from Sheet 1 & Sheet 2');
        this.createPairedTables(data, sheet1Data.data, sheet2Data.data, 'Sheet 1', 'Sheet 2');
      } else {
        // Only Sheet 1 available - create single-sheet tables
        console.log('Creating single-sheet tables from Sheet 1 only (Sheet 2 is empty/missing)');
        this.createSingleSheetTables(data, sheet1Data.data, 'Sheet 1');
      }
    } else if (sheet2Data && sheet2Data.hasData) {
      // Only Sheet 2 available - create single-sheet tables
      console.log('Creating single-sheet tables from Sheet 2 only (Sheet 1 is empty/missing)');
      this.createSingleSheetTables(data, sheet2Data.data, 'Sheet 2');
    }

    // Process Sheet 3 & 4 pair
    if (sheet3Data && sheet3Data.hasData) {
      if (sheet4Data && sheet4Data.hasData) {
        // Both sheets available - create paired tables
        console.log('Creating paired tables from Sheet 3 & Sheet 4');
        this.createPairedTables(data, sheet3Data.data, sheet4Data.data, 'Sheet 3', 'Sheet 4');
      } else {
        // Only Sheet 3 available - create single-sheet tables
        console.log('Creating single-sheet tables from Sheet 3 only (Sheet 4 is empty/missing)');
        this.createSingleSheetTables(data, sheet3Data.data, 'Sheet 3');
      }
    } else if (sheet4Data && sheet4Data.hasData) {
      // Only Sheet 4 available - create single-sheet tables
      console.log('Creating single-sheet tables from Sheet 4 only (Sheet 3 is empty/missing)');
      this.createSingleSheetTables(data, sheet4Data.data, 'Sheet 4');
    }

    console.log(`Generated ${data.tables.length} tables total`);
    
    if (data.tables.length === 0) {
      throw new Error('No tables could be generated from the available data. Please check your Excel file format.');
    }

    return data;
  }

  // Check if sheet data contains meaningful information
  hasValidData(sheetData) {
    if (!sheetData || !sheetData.headers || !sheetData.rows) {
      return false;
    }
    
    // Check if there are at least 3 columns (including the first 2 that we skip)
    if (sheetData.headers.length < 3) {
      return false;
    }
    
    // Check if there are data rows
    if (sheetData.rows.length === 0) {
      return false;
    }
    
    // Check if there's actual data in columns beyond the first 2
    let hasDataInColumns = false;
    for (let colIndex = 2; colIndex < sheetData.headers.length; colIndex++) {
      const columnValues = this.getColumnValues(sheetData.rows, colIndex);
      if (Object.keys(columnValues).length > 0) {
        hasDataInColumns = true;
        break;
      }
    }
    
    return hasDataInColumns;
  }

  // Create tables from paired sheets (original logic)
  createPairedTables(data, sheet1Data, sheet2Data, sheet1Name, sheet2Name) {
    // Validate that sheets have same number of columns
    if (sheet1Data.headers.length !== sheet2Data.headers.length) {
      console.warn(`${sheet1Name} and ${sheet2Name} have different number of columns. Using available columns.`);
    }

    const maxColumns = Math.max(sheet1Data.headers.length, sheet2Data.headers.length);
    const minColumns = Math.min(sheet1Data.headers.length, sheet2Data.headers.length);

    // Generate tables for each column (excluding first 2 columns)
    for (let colIndex = 2; colIndex < minColumns; colIndex++) {
      const header = sheet1Data.headers[colIndex] || sheet2Data.headers[colIndex] || `Column ${colIndex + 1}`;
      
      const sheet1Values = this.getColumnValues(sheet1Data.rows, colIndex);
      const sheet2Values = this.getColumnValues(sheet2Data.rows, colIndex);
      
      // Only create table if at least one sheet has data for this column
      if (Object.keys(sheet1Values).length > 0 || Object.keys(sheet2Values).length > 0) {
        const table = this.createTable(header, sheet1Values, sheet2Values, sheet1Name, sheet2Name);
        data.tables.push(table);
      }
    }
  }

  // Create tables from single sheet (new logic)
  createSingleSheetTables(data, sheetData, sheetName) {
    // Generate tables for each column (excluding first 2 columns)
    for (let colIndex = 2; colIndex < sheetData.headers.length; colIndex++) {
      const header = sheetData.headers[colIndex];
      
      const sheetValues = this.getColumnValues(sheetData.rows, colIndex);
      
      // Only create table if this column has data
      if (Object.keys(sheetValues).length > 0) {
        const table = this.createSingleSheetTable(header, sheetValues, sheetName);
        data.tables.push(table);
      }
    }
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

  // Updated createTable method to handle both paired and single sheet scenarios
  createTable(header, trialValues, controlValues, trialGroupName = 'Trial Group', controlGroupName = 'Control Group') {
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
      rows: rows,
      sourceInfo: {
        trialGroup: trialGroupName,
        controlGroup: controlGroupName
      },
      isPaired: controlTotal > 0  // Flag to indicate if this is paired data or single sheet
    };
  }

  // New method to create single-sheet tables
  createSingleSheetTable(header, sheetValues, sheetName) {
    const rows = [];
    let total = 0;

    // Create rows for each unique value
    Object.keys(sheetValues).sort().forEach(value => {
      const count = sheetValues[value];
      total += count;

      rows.push({
        category: value,
        trialGroup: count,  // Use the single sheet data as trial group
        controlGroup: 0,    // No control group data
        total: count,
        percentage: 0 // Will be calculated after we have the grand total
      });
    });

    // Calculate percentages
    rows.forEach(row => {
      row.percentage = total > 0 ? ((row.total / total) * 100).toFixed(2) : '0.00';
    });

    // Add total row
    rows.push({
      category: 'Total',
      trialGroup: total,
      controlGroup: 0,
      total: total,
      percentage: '100.00'
    });

    return {
      title: header,
      rows: rows,
      sourceInfo: {
        trialGroup: sheetName,
        controlGroup: 'N/A (Single Sheet)'
      },
      isPaired: false  // Flag to indicate this is single sheet data
    };
  }
}

module.exports = new ExcelService();