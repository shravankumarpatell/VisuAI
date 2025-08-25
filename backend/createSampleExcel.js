const ExcelJS = require('exceljs');

async function createSampleExcel() {
  const workbook = new ExcelJS.Workbook();
  
  // Create Trial Group sheet
  const trialSheet = workbook.addWorksheet('Trial Group');
  
  // Add headers
  trialSheet.columns = [
    { header: 'Sl.No', key: 'slno', width: 10 },
    { header: 'OPD No', key: 'opdno', width: 15 },
    { header: 'Age', key: 'age', width: 15 },
    { header: 'Gender', key: 'gender', width: 15 },
    { header: 'Religion', key: 'religion', width: 15 },
    { header: 'Occupation', key: 'occupation', width: 20 },
    { header: 'Diet', key: 'diet', width: 15 },
    { header: 'Habitat', key: 'habitat', width: 15 },
    { header: 'Sleep', key: 'sleep', width: 15 },
    { header: 'Socio-Economic Status', key: 'ses', width: 25 },
    { header: 'Education', key: 'education', width: 20 },
    { header: 'Marital Status', key: 'marital', width: 20 }
  ];

  // Style headers
  trialSheet.getRow(1).font = { bold: true };
  trialSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
  };

  // Add sample data for Trial Group
  const trialData = [];
  for (let i = 1; i <= 20; i++) {
    trialData.push({
      slno: i,
      opdno: `OPD${1000 + i}`,
      age: ['20-30', '31-40', '41-50', '51-60'][Math.floor(Math.random() * 4)],
      gender: ['Male', 'Female'][Math.floor(Math.random() * 2)],
      religion: ['Hindu', 'Muslim', 'Christian'][Math.floor(Math.random() * 3)],
      occupation: ['Labour', 'Housewife', 'Business', 'Student', 'Teacher', 'Farmer', 'Driver'][Math.floor(Math.random() * 7)],
      diet: ['Veg', 'Mixed'][Math.floor(Math.random() * 2)],
      habitat: ['Rural', 'Urban'][Math.floor(Math.random() * 2)],
      sleep: ['Good', 'Disturbed'][Math.random() > 0.2 ? 0 : 1],
      ses: ['Lower Class', 'Middle Class'][Math.floor(Math.random() * 2)],
      education: ['Primary', 'SSLC', 'PUC', 'Graduate', 'Post Graduate'][Math.floor(Math.random() * 5)],
      marital: ['Married', 'Unmarried'][Math.random() > 0.3 ? 0 : 1]
    });
  }
  trialSheet.addRows(trialData);

  // Create Control Group sheet
  const controlSheet = workbook.addWorksheet('Control Group');
  
  // Copy headers
  controlSheet.columns = trialSheet.columns;
  controlSheet.getRow(1).font = { bold: true };
  controlSheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE0E0E0' }
  };

  // Add sample data for Control Group
  const controlData = [];
  for (let i = 1; i <= 20; i++) {
    controlData.push({
      slno: i,
      opdno: `OPD${2000 + i}`,
      age: ['20-30', '31-40', '41-50', '51-60'][Math.floor(Math.random() * 4)],
      gender: ['Male', 'Female'][Math.floor(Math.random() * 2)],
      religion: ['Hindu', 'Muslim', 'Christian'][Math.floor(Math.random() * 3)],
      occupation: ['Labour', 'Housewife', 'Business', 'Student', 'Teacher', 'Farmer', 'Driver'][Math.floor(Math.random() * 7)],
      diet: ['Veg', 'Mixed'][Math.floor(Math.random() * 2)],
      habitat: ['Rural', 'Urban'][Math.floor(Math.random() * 2)],
      sleep: ['Good', 'Disturbed'][Math.random() > 0.2 ? 0 : 1],
      ses: ['Lower Class', 'Middle Class'][Math.floor(Math.random() * 2)],
      education: ['Primary', 'SSLC', 'PUC', 'Graduate', 'Post Graduate'][Math.floor(Math.random() * 5)],
      marital: ['Married', 'Unmarried'][Math.random() > 0.3 ? 0 : 1]
    });
  }
  controlSheet.addRows(controlData);

  // Save the workbook
  await workbook.xlsx.writeFile('sample-masterchart.xlsx');
  console.log('Sample Excel file created: sample-masterchart.xlsx');
}

// Run the function
createSampleExcel().catch(console.error);