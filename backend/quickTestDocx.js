
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
