// const ExcelJS = require('exceljs');

// class ExcelService {
//   async readExcelFile(filePath) {
//     const workbook = new ExcelJS.Workbook();
//     await workbook.xlsx.readFile(filePath);
    
//     const data = {
//       trialGroup: null,
//       controlGroup: null,
//       tables: []
//     };

//     // Find sheets for Trial Group and Control Group
//     const trialSheet = this.findSheet(workbook, ['trial', 'trial group']);
//     const controlSheet = this.findSheet(workbook, ['control', 'control group']);

//     if (!trialSheet || !controlSheet) {
//       throw new Error('Could not find Trial Group or Control Group sheets in the Excel file');
//     }

//     // Read data from both sheets
//     const trialData = this.extractSheetData(trialSheet);
//     const controlData = this.extractSheetData(controlSheet);

//     if (trialData.headers.length !== controlData.headers.length) {
//       throw new Error('Trial and Control groups have different number of columns');
//     }

//     // Generate tables for each column (excluding first 2)
//     for (let colIndex = 2; colIndex < trialData.headers.length; colIndex++) {
//       const header = trialData.headers[colIndex];
      
//       // Get unique values and their counts
//       const trialValues = this.getColumnValues(trialData.rows, colIndex);
//       const controlValues = this.getColumnValues(controlData.rows, colIndex);
      
//       const table = this.createTable(header, trialValues, controlValues);
//       data.tables.push(table);
//     }

//     return data;
//   }

//   findSheet(workbook, sheetNames) {
//     for (const worksheet of workbook.worksheets) {
//       const sheetName = worksheet.name.toLowerCase();
//       if (sheetNames.some(name => sheetName.includes(name))) {
//         return worksheet;
//       }
//     }
//     return null;
//   }

//   extractSheetData(worksheet) {
//     const data = {
//       headers: [],
//       rows: []
//     };

//     let headerRow = null;
//     let dataStartRow = null;

//     // Find header row (usually the first non-empty row)
//     worksheet.eachRow((row, rowNumber) => {
//       if (!headerRow && row.actualCellCount > 0) {
//         headerRow = rowNumber;
//         row.eachCell((cell, colNumber) => {
//           data.headers[colNumber - 1] = cell.value ? cell.value.toString() : '';
//         });
//       } else if (headerRow && !dataStartRow) {
//         dataStartRow = rowNumber;
//       }
//     });

//     // Extract data rows
//     worksheet.eachRow((row, rowNumber) => {
//       if (rowNumber >= dataStartRow) {
//         const rowData = [];
//         row.eachCell((cell, colNumber) => {
//           rowData[colNumber - 1] = cell.value;
//         });
//         if (rowData.some(val => val !== null && val !== undefined && val !== '')) {
//           data.rows.push(rowData);
//         }
//       }
//     });

//     return data;
//   }

//   getColumnValues(rows, colIndex) {
//     const values = {};
    
//     rows.forEach(row => {
//       const value = row[colIndex];
//       if (value !== null && value !== undefined && value !== '') {
//         const key = value.toString();
//         values[key] = (values[key] || 0) + 1;
//       }
//     });

//     return values;
//   }

//   createTable(header, trialValues, controlValues) {
//     // Get all unique values
//     const allValues = new Set([
//       ...Object.keys(trialValues),
//       ...Object.keys(controlValues)
//     ]);

//     const rows = [];
//     let trialTotal = 0;
//     let controlTotal = 0;

//     // Create rows for each unique value
//     allValues.forEach(value => {
//       const trialCount = trialValues[value] || 0;
//       const controlCount = controlValues[value] || 0;
//       const total = trialCount + controlCount;
      
//       trialTotal += trialCount;
//       controlTotal += controlCount;

//       rows.push({
//         category: value,
//         trialGroup: trialCount,
//         controlGroup: controlCount,
//         total: total,
//         percentage: 0 // Will be calculated after we have the grand total
//       });
//     });

//     const grandTotal = trialTotal + controlTotal;

//     // Calculate percentages and add total row
//     rows.forEach(row => {
//       row.percentage = grandTotal > 0 ? ((row.total / grandTotal) * 100).toFixed(2) : '0.00';
//     });

//     // Add total row
//     rows.push({
//       category: 'Total',
//       trialGroup: trialTotal,
//       controlGroup: controlTotal,
//       total: grandTotal,
//       percentage: '100.00'
//     });

//     return {
//       title: header,
//       rows: rows
//     };
//   }
// }

// module.exports = new ExcelService();



const ExcelJS = require('exceljs');

class ExcelService {
  async readExcelFile(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const data = {
      trialGroup: null,
      controlGroup: null,
      tables: []
    };

    // Find sheets for Trial Group and Control Group
    const trialSheet = this.findSheet(workbook, ['trial', 'trial group']);
    const controlSheet = this.findSheet(workbook, ['control', 'control group']);

    // If not found by name, try by position (first two sheets)
    let usingPositionFallback = false;
    if (!trialSheet || !controlSheet) {
      console.log('Could not find sheets by name, trying by position...');
      const sheet1 = this.findSheetByPosition(workbook, 0);
      const sheet2 = this.findSheetByPosition(workbook, 1);
      
      if (sheet1 && sheet2) {
        console.log(`Using sheets by position: "${sheet1.name}" as Trial Group, "${sheet2.name}" as Control Group`);
        data.trialGroup = sheet1;
        data.controlGroup = sheet2;
        usingPositionFallback = true;
      }
    }

    if (!usingPositionFallback && (!trialSheet || !controlSheet)) {
      // Provide helpful error message with actual sheet names
      const actualSheetNames = workbook.worksheets.map(ws => ws.name).join(', ');
      throw new Error(
        `Could not find required sheets. Expected sheets containing "Trial Group" and "Control Group". ` +
        `Found sheets: ${actualSheetNames}. ` +
        `Please rename your sheets or ensure the first sheet is Trial Group and second is Control Group.`
      );
    }

    // Read data from both sheets
    const trialData = this.extractSheetData(usingPositionFallback ? data.trialGroup : trialSheet);
    const controlData = this.extractSheetData(usingPositionFallback ? data.controlGroup : controlSheet);

    if (trialData.headers.length !== controlData.headers.length) {
      throw new Error('Trial and Control groups have different number of columns');
    }

    // Generate tables for each column (excluding first 2)
    for (let colIndex = 2; colIndex < trialData.headers.length; colIndex++) {
      const header = trialData.headers[colIndex];
      
      // Get unique values and their counts
      const trialValues = this.getColumnValues(trialData.rows, colIndex);
      const controlValues = this.getColumnValues(controlData.rows, colIndex);
      
      const table = this.createTable(header, trialValues, controlValues);
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
        const key = value.toString();
        values[key] = (values[key] || 0) + 1;
      }
    });

    return values;
  }

  createTable(header, trialValues, controlValues) {
    // Get all unique values
    const allValues = new Set([
      ...Object.keys(trialValues),
      ...Object.keys(controlValues)
    ]);

    const rows = [];
    let trialTotal = 0;
    let controlTotal = 0;

    // Create rows for each unique value
    allValues.forEach(value => {
      const trialCount = trialValues[value] || 0;
      const controlCount = controlValues[value] || 0;
      const total = trialCount + controlCount;
      
      trialTotal += trialCount;
      controlTotal += controlCount;

      rows.push({
        category: value,
        trialGroup: trialCount,
        controlGroup: controlCount,
        total: total,
        percentage: 0 // Will be calculated after we have the grand total
      });
    });

    const grandTotal = trialTotal + controlTotal;

    // Calculate percentages and add total row
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