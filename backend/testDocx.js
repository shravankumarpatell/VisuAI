const { Document, Packer, Paragraph, TextRun } = require('docx');
const fs = require('fs').promises;

async function testDocx() {
  console.log('Testing docx library...\n');
  
  try {
    // Create a simple document
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: "Test Document",
                bold: true,
                size: 32
              })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun("This is a test to verify docx library is working correctly.")
            ]
          })
        ]
      }]
    });

    // Try to save it
    console.log('Creating document buffer...');
    const buffer = await Packer.toBuffer(doc);
    console.log('✅ Buffer created successfully');
    
    // Write to file
    console.log('Writing to file...');
    await fs.writeFile('test-document.docx', buffer);
    console.log('✅ File written successfully: test-document.docx');
    
    // Show buffer size
    console.log(`File size: ${(buffer.length / 1024).toFixed(2)} KB`);
    
    console.log('\n✅ docx library is working correctly!');
    console.log('\nThe issue might be with:');
    console.log('1. The specific version of docx installed');
    console.log('2. Missing Packer import in wordService.js');
    console.log('\nMake sure wordService.js imports Packer:');
    console.log('const { Document, Packer, Paragraph, ... } = require("docx");');
    
  } catch (error) {
    console.error('❌ Error:', error.message);
    console.error('\nTroubleshooting steps:');
    console.error('1. Check docx version: npm list docx');
    console.error('2. Reinstall docx: npm uninstall docx && npm install docx');
    console.error('3. Clear npm cache: npm cache clean --force');
  }
}

// Check docx version
try {
  const docxPackage = require('docx/package.json');
  console.log(`docx version: ${docxPackage.version}`);
} catch (e) {
  console.log('Could not read docx version');
}

testDocx();