const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

console.log(`
╔════════════════════════════════════════════╗
║     Thesis Helper Troubleshooting Tool     ║
╚════════════════════════════════════════════╝
`);

// Check Node.js version
console.log('🔍 Checking environment...\n');
console.log(`✓ Node.js version: ${process.version}`);
console.log(`✓ Platform: ${process.platform}`);
console.log(`✓ Current directory: ${process.cwd()}`);

// Check if dependencies are installed
console.log('\n🔍 Checking dependencies...\n');
const dependencies = [
  'express',
  'cors',
  'multer',
  'exceljs',
  'mammoth',
  'pdfkit',
  'chart.js',
  'canvas',
  'docx'
];

let missingDeps = [];
dependencies.forEach(dep => {
  try {
    require.resolve(dep);
    console.log(`✅ ${dep} - Installed`);
  } catch (e) {
    console.log(`❌ ${dep} - Not installed`);
    missingDeps.push(dep);
  }
});

if (missingDeps.length > 0) {
  console.log('\n⚠️  Missing dependencies detected!');
  console.log('Run: npm install');
}

// Check uploads directory
console.log('\n🔍 Checking uploads directory...\n');
const uploadsDir = path.join(__dirname, 'uploads');
if (fs.existsSync(uploadsDir)) {
  console.log('✅ Uploads directory exists');
  const files = fs.readdirSync(uploadsDir);
  console.log(`   Files in uploads: ${files.length}`);
  
  // Check directory permissions
  try {
    const testFile = path.join(uploadsDir, '.test');
    fs.writeFileSync(testFile, 'test');
    fs.unlinkSync(testFile);
    console.log('✅ Write permissions OK');
  } catch (e) {
    console.log('❌ No write permissions to uploads directory');
  }
} else {
  console.log('❌ Uploads directory missing');
  console.log('   Creating uploads directory...');
  try {
    fs.mkdirSync(uploadsDir, { recursive: true });
    console.log('   ✅ Directory created');
  } catch (e) {
    console.log('   ❌ Failed to create directory:', e.message);
  }
}

// Test Excel file processing
console.log('\n🔍 Testing Excel functionality...\n');
const testExcelFile = path.join(__dirname, 'sample-masterchart.xlsx');

if (fs.existsSync(testExcelFile)) {
  console.log('✅ Sample Excel file found');
  
  // Try to read it
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx.readFile(testExcelFile)
    .then(() => {
      console.log('✅ Excel file reading works');
      console.log(`   Sheets found: ${workbook.worksheets.map(ws => ws.name).join(', ')}`);
      
      // Test the specific sheet lookup
      const sheetNames = workbook.worksheets.map(ws => ws.name.toLowerCase());
      const hasTrialGroup = sheetNames.some(name => 
        name.includes('trial') && name.includes('group')
      );
      const hasControlGroup = sheetNames.some(name => 
        name.includes('control') && name.includes('group')
      );
      
      if (hasTrialGroup && hasControlGroup) {
        console.log('✅ Required sheets found (Trial Group & Control Group)');
      } else {
        console.log('⚠️  Required sheets not found by name');
        console.log('   The system will use the first two sheets as fallback');
      }
      
      // Common issues summary
      console.log('\n📋 Common Issues & Solutions:\n');
      console.log('1. "Could not find Trial Group or Control Group sheets"');
      console.log('   → Rename your Excel sheets to "Trial Group" and "Control Group"');
      console.log('   → Or ensure you have at least 2 sheets (they\'ll be used automatically)');
      console.log('\n2. "File upload failed"');
      console.log('   → Check file size (must be under 10MB)');
      console.log('   → Ensure file is .xlsx or .xls format');
      console.log('\n3. "Canvas installation failed"');
      console.log('   → Install system dependencies (see README)');
      console.log('   → Try: npm rebuild canvas');
      
      console.log('\n✅ Troubleshooting complete!\n');
    })
    .catch(err => {
      console.log('❌ Error reading Excel file:', err.message);
    });
} else {
  console.log('❌ Sample Excel file not found');
  console.log('   Run: node createSampleExcel.js');
}

// Show next steps
console.log('\n📌 Next Steps:\n');
console.log('1. If you see any ❌ above, fix those issues first');
console.log('2. Generate a sample file: node createSampleExcel.js');
console.log('3. Check your Excel file: node checkExcelFile.js your-file.xlsx');
console.log('4. Start the server: npm run dev');
console.log('5. Try uploading the sample file first\n');