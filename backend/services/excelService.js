
// const ExcelJS = require('exceljs');

// class ExcelService {
//   async readExcelFile(filePath) {
//     const workbook = new ExcelJS.Workbook();
//     await workbook.xlsx.readFile(filePath);
    
// const data = {
//   trialGroup: null,
//   controlGroup: null,
//   thirdGroup: null,    // ADD THIS
//   fourthGroup: null,   // ADD THIS
//   tables: []
// };

// // Find sheets for Trial Group and Control Group
// const trialSheet = this.findSheet(workbook, ['trial', 'trial group']);
// const controlSheet = this.findSheet(workbook, ['control', 'control group']);
// const thirdSheet = this.findSheetByPosition(workbook, 2);   // ADD THIS
// const fourthSheet = this.findSheetByPosition(workbook, 3);  // ADD THIS

//     // If not found by name, try by position (first two sheets)
//     // If not found by name, try by position (first four sheets)
// let usingPositionFallback = false;
// if (!trialSheet || !controlSheet) {
//   console.log('Could not find sheets by name, trying by position...');
//   const sheet1 = this.findSheetByPosition(workbook, 0);
//   const sheet2 = this.findSheetByPosition(workbook, 1);
//   const sheet3 = this.findSheetByPosition(workbook, 2);
//   const sheet4 = this.findSheetByPosition(workbook, 3);
  
//   if (sheet1 && sheet2 && sheet3 && sheet4) {  // CHANGE: require all 4 sheets
//     console.log(`Using sheets by position: "${sheet1.name}", "${sheet2.name}", "${sheet3.name}", "${sheet4.name}"`);
//     data.trialGroup = sheet1;
//     data.controlGroup = sheet2;
//     data.thirdGroup = sheet3;   // ADD THIS
//     data.fourthGroup = sheet4;  // ADD THIS
//     usingPositionFallback = true;
//   }
// }

//    if (!usingPositionFallback && (!trialSheet || !controlSheet || !thirdSheet || !fourthSheet)) {
//   const actualSheetNames = workbook.worksheets.map(ws => ws.name).join(', ');
//   throw new Error(
//     `Could not find required sheets. Expected 4 sheets. ` +
//     `Found sheets: ${actualSheetNames}. ` +
//     `Please ensure you have at least 4 sheets in your Excel file.`
//   );
// }

//     // Read data from both sheets
// // Read data from all four sheets
// const trialData = this.extractSheetData(usingPositionFallback ? data.trialGroup : trialSheet);
// const controlData = this.extractSheetData(usingPositionFallback ? data.controlGroup : controlSheet);
// const thirdData = this.extractSheetData(usingPositionFallback ? data.thirdGroup : thirdSheet);   // ADD THIS
// const fourthData = this.extractSheetData(usingPositionFallback ? data.fourthGroup : fourthSheet); // ADD THIS

// if (trialData.headers.length !== controlData.headers.length || 
//     trialData.headers.length !== thirdData.headers.length || 
//     trialData.headers.length !== fourthData.headers.length) {
//   throw new Error('All sheets must have the same number of columns');
// }

// // Generate tables for each column (excluding first 2) from all 4 sheets
// for (let colIndex = 2; colIndex < trialData.headers.length; colIndex++) {
//   const header = trialData.headers[colIndex];
  
//   // Get unique values and their counts from all 4 sheets
//   const trialValues = this.getColumnValues(trialData.rows, colIndex);
//   const controlValues = this.getColumnValues(controlData.rows, colIndex);
//   const thirdValues = this.getColumnValues(thirdData.rows, colIndex);     // ADD THIS
//   const fourthValues = this.getColumnValues(fourthData.rows, colIndex);   // ADD THIS
  
//   const table = this.createTable(header, trialValues, controlValues, thirdValues, fourthValues); // UPDATE THIS
//   data.tables.push(table);
// }

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

//   // Alternative method to find sheets by position if names don't match
//   findSheetByPosition(workbook, position) {
//     if (workbook.worksheets && workbook.worksheets.length > position) {
//       return workbook.worksheets[position];
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

// createTable(header, trialValues, controlValues, thirdValues, fourthValues) {  // UPDATE SIGNATURE
//   // Get all unique values from all 4 groups
//   const allValues = new Set([
//     ...Object.keys(trialValues),
//     ...Object.keys(controlValues),
//     ...Object.keys(thirdValues),    // ADD THIS
//     ...Object.keys(fourthValues)    // ADD THIS
//   ]);

//   const rows = [];
//   let trialTotal = 0;
//   let controlTotal = 0;
//   let thirdTotal = 0;    // ADD THIS
//   let fourthTotal = 0;   // ADD THIS

//   // Create rows for each unique value
//   allValues.forEach(value => {
//     const trialCount = trialValues[value] || 0;
//     const controlCount = controlValues[value] || 0;
//     const thirdCount = thirdValues[value] || 0;     // ADD THIS
//     const fourthCount = fourthValues[value] || 0;   // ADD THIS
//     const total = trialCount + controlCount + thirdCount + fourthCount; // UPDATE THIS
    
//     trialTotal += trialCount;
//     controlTotal += controlCount;
//     thirdTotal += thirdCount;     // ADD THIS
//     fourthTotal += fourthCount;   // ADD THIS

//     rows.push({
//       category: value,
//       trialGroup: trialCount,
//       controlGroup: controlCount,
//       thirdGroup: thirdCount,     // ADD THIS
//       fourthGroup: fourthCount,   // ADD THIS
//       total: total,
//       percentage: 0
//     });
//   });

//   const grandTotal = trialTotal + controlTotal + thirdTotal + fourthTotal; // UPDATE THIS

//   // Calculate percentages and add total row
//   rows.forEach(row => {
//     row.percentage = grandTotal > 0 ? ((row.total / grandTotal) * 100).toFixed(2) : '0.00';
//   });

//   // Add total row
//   rows.push({
//     category: 'Total',
//     trialGroup: trialTotal,
//     controlGroup: controlTotal,
//     thirdGroup: thirdTotal,     // ADD THIS
//     fourthGroup: fourthTotal,   // ADD THIS
//     total: grandTotal,
//     percentage: '100.00'
//   });

//   return {
//     title: header,
//     rows: rows
//   };
// }
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
      
      const table = this.createTable(`${header} (Sheets 1-2)`, sheet1Values, sheet2Values);
      data.tables.push(table);
    }

    // Generate tables for Sheet 3 & Sheet 4 pair (excluding first 2 columns)
    for (let colIndex = 2; colIndex < sheet3Data.headers.length; colIndex++) {
      const header = sheet3Data.headers[colIndex];
      
      const sheet3Values = this.getColumnValues(sheet3Data.rows, colIndex);
      const sheet4Values = this.getColumnValues(sheet4Data.rows, colIndex);
      
      const table = this.createTable(`${header} (Sheets 3-4)`, sheet3Values, sheet4Values);
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