const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log(`
╔═══════════════════════════════════════════╗
║     Fix docx Library Issue                ║
╚═══════════════════════════════════════════╝
`);

console.log('🔧 Fixing the docx library issue...\n');

// Step 1: Check current docx version
console.log('1️⃣ Checking current docx version...');
try {
  const docxVersion = execSync('npm list docx --depth=0', { encoding: 'utf-8' });
  console.log(docxVersion);
} catch (e) {
  console.log('docx not found or error checking version');
}

// Step 2: Update package.json with specific docx version
console.log('\n2️⃣ Updating package.json with correct docx version...');
const packagePath = path.join(__dirname, 'package.json');
try {
  const packageJson = JSON.parse(fs.readFileSync(packagePath, 'utf-8'));
  packageJson.dependencies.docx = "^8.5.0";  // Specific working version
  fs.writeFileSync(packagePath, JSON.stringify(packageJson, null, 2));
  console.log('✅ package.json updated');
} catch (e) {
  console.error('❌ Error updating package.json:', e.message);
}

// Step 3: Create backup of current wordService.js
console.log('\n3️⃣ Creating backup of wordService.js...');
const wordServicePath = path.join(__dirname, 'services', 'wordService.js');
const backupPath = path.join(__dirname, 'services', 'wordService.backup.js');
try {
  fs.copyFileSync(wordServicePath, backupPath);
  console.log('✅ Backup created: wordService.backup.js');
} catch (e) {
  console.error('❌ Error creating backup:', e.message);
}

// Step 4: Replace wordService.js with fixed version
console.log('\n4️⃣ Updating wordService.js with fixed version...');
try {
  const fixedServicePath = path.join(__dirname, 'services', 'wordServiceFixed.js');
  if (fs.existsSync(fixedServicePath)) {
    fs.copyFileSync(fixedServicePath, wordServicePath);
    console.log('✅ wordService.js updated with fixed version');
  } else {
    console.log('⚠️  wordServiceFixed.js not found, applying inline fix...');
    
    // Read current file
    let content = fs.readFileSync(wordServicePath, 'utf-8');
    
    // Fix imports
    if (!content.includes('Packer')) {
      content = content.replace(
        'const { Document,',
        'const { Document, Packer,'
      );
    }
    
    // Fix save method
    content = content.replace(
      'const buffer = await doc.save();',
      'const buffer = await Packer.toBuffer(doc);'
    );
    
    // Fix text conversions
    content = content.replace(/text: row\.category,/g, 'text: String(row.category),');
    content = content.replace(/text: row\.trialGroup\.toString\(\),/g, 'text: String(row.trialGroup),');
    content = content.replace(/text: row\.controlGroup\.toString\(\),/g, 'text: String(row.controlGroup),');
    content = content.replace(/text: row\.total\.toString\(\),/g, 'text: String(row.total),');
    
    fs.writeFileSync(wordServicePath, content);
    console.log('✅ wordService.js fixed inline');
  }
} catch (e) {
  console.error('❌ Error updating wordService.js:', e.message);
}

// Step 5: Reinstall docx
console.log('\n5️⃣ Reinstalling docx library...');
console.log('Running: npm uninstall docx');
try {
  execSync('npm uninstall docx', { stdio: 'inherit' });
} catch (e) {
  console.log('Warning: Error uninstalling docx');
}

console.log('\nRunning: npm install docx@^8.5.0');
try {
  execSync('npm install docx@^8.5.0', { stdio: 'inherit' });
  console.log('✅ docx installed successfully');
} catch (e) {
  console.error('❌ Error installing docx:', e.message);
  console.log('\nTry running manually:');
  console.log('npm cache clean --force');
  console.log('npm install docx@^8.5.0');
}

// Step 6: Test the fix
console.log('\n6️⃣ Testing the fix...');
console.log('Run: node testDocx.js');

console.log(`
✅ Fix applied! Next steps:
1. Restart your server: npm run dev
2. Try uploading your Excel file again
3. If still having issues, run: node testDocx.js

If the problem persists:
- Check the error logs
- Make sure all dependencies are installed: npm install
- Try clearing npm cache: npm cache clean --force
`);

// Create a simple test to verify
const testCode = `
const { Document, Packer, Paragraph, TextRun } = require('docx');

async function quickTest() {
  try {
    const doc = new Document({
      sections: [{
        children: [new Paragraph({ children: [new TextRun("Test")] })]
      }]
    });
    const buffer = await Packer.toBuffer(doc);
    console.log('✅ Quick test passed! docx is working.');
  } catch (e) {
    console.error('❌ Quick test failed:', e.message);
  }
}
quickTest();
`;

fs.writeFileSync(path.join(__dirname, 'quickTestDocx.js'), testCode);
console.log('\nCreated quickTestDocx.js - Run it to verify the fix.');