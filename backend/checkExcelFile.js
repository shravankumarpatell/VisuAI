const ExcelJS = require('exceljs');
const path = require('path');

async function checkExcelFile(filePath) {
  console.log('\n📋 Excel File Structure Check\n');
  console.log('File:', path.basename(filePath));
  console.log('=' .repeat(50));
  
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    console.log(`\n✅ File loaded successfully`);
    console.log(`📊 Number of sheets: ${workbook.worksheets.length}`);
    
    // Check each sheet
    workbook.worksheets.forEach((worksheet, index) => {
      console.log(`\n📄 Sheet ${index + 1}: "${worksheet.name}"`);
      console.log(`   - Rows: ${worksheet.rowCount}`);
      console.log(`   - Columns: ${worksheet.columnCount}`);
      
      // Check headers (first row)
      if (worksheet.rowCount > 0) {
        const headers = [];
        worksheet.getRow(1).eachCell((cell, colNumber) => {
          headers.push(cell.value || `(empty col ${colNumber})`);
        });
        console.log(`   - Headers: ${headers.join(', ')}`);
      }
      
      // Check data rows
      let dataRowCount = 0;
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          const hasData = row.values.some(val => val !== null && val !== undefined && val !== '');
          if (hasData) dataRowCount++;
        }
      });
      console.log(`   - Data rows: ${dataRowCount}`);
    });
    
    // Check for required sheets
    console.log('\n🔍 Sheet Name Analysis:');
    const sheetNames = workbook.worksheets.map(ws => ws.name);
    
    const hasTrialGroup = sheetNames.some(name => 
      name.toLowerCase().includes('trial') && name.toLowerCase().includes('group')
    );
    const hasControlGroup = sheetNames.some(name => 
      name.toLowerCase().includes('control') && name.toLowerCase().includes('group')
    );
    
    console.log(`   - Has "Trial Group" sheet: ${hasTrialGroup ? '✅ Yes' : '❌ No'}`);
    console.log(`   - Has "Control Group" sheet: ${hasControlGroup ? '✅ Yes' : '❌ No'}`);
    
    // Recommendations
    console.log('\n💡 Recommendations:');
    if (!hasTrialGroup || !hasControlGroup) {
      console.log('   ⚠️  Your sheets need to be renamed to include "Trial Group" and "Control Group"');
      console.log('   ⚠️  OR ensure your first sheet is Trial data and second sheet is Control data');
      
      if (workbook.worksheets.length >= 2) {
        console.log(`\n   📝 Quick fix - Rename your sheets:`);
        console.log(`      - Rename "${sheetNames[0]}" to "Trial Group"`);
        console.log(`      - Rename "${sheetNames[1]}" to "Control Group"`);
      }
    } else {
      console.log('   ✅ Sheet names are correctly formatted!');
    }
    
    // Check column structure
    if (workbook.worksheets.length > 0) {
      const firstSheet = workbook.worksheets[0];
      if (firstSheet.columnCount < 3) {
        console.log('   ⚠️  Your sheet has less than 3 columns. Ensure you have at least:');
        console.log('      Column 1: Sl.No');
        console.log('      Column 2: OPD No');
        console.log('      Column 3+: Your data columns (Age, Gender, etc.)');
      }
    }
    
    console.log('\n' + '=' .repeat(50));
    console.log('Check complete!\n');
    
  } catch (error) {
    console.error('❌ Error reading file:', error.message);
    console.log('\nMake sure:');
    console.log('1. The file exists at the specified path');
    console.log('2. The file is a valid Excel file (.xlsx or .xls)');
    console.log('3. The file is not corrupted or password protected');
  }
}

// Check if file path is provided as command line argument
const filePath = process.argv[2];

if (!filePath) {
  console.log('Usage: node checkExcelFile.js <path-to-excel-file>');
  console.log('Example: node checkExcelFile.js sample-masterchart.xlsx');
} else {
  checkExcelFile(filePath);
}