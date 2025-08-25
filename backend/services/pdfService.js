const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

class PDFService {
  async createPDF(graphs, originalFilename) {
    return new Promise((resolve, reject) => {
      try {
        const doc = new PDFDocument({
          size: 'A4',
          layout: 'portrait',
          margin: 50,
          info: {
            Title: 'Thesis Statistical Analysis Graphs',
            Author: 'Thesis Helper',
            Subject: 'Statistical Analysis',
            Keywords: 'thesis, graphs, statistical analysis'
          },
          compress: false  // No compression for better quality
        });

        const timestamp = Date.now();
        const outputPath = path.join('uploads', `graphs-${timestamp}.pdf`);
        const stream = fs.createWriteStream(outputPath);

        doc.pipe(stream);

        // Add title page with better styling
        doc.fontSize(28)
           .font('Helvetica-Bold')
           .fillColor('#1f2937')
           .text('Sample Size Estimation', { align: 'center' });
        
        doc.moveDown();
        doc.fontSize(22)
           .fillColor('#374151')
           .text('Statistical Analysis Graphs', { align: 'center' });
        
        doc.moveDown(2);
        doc.fontSize(14)
           .font('Helvetica')
           .fillColor('#6b7280')
           .text(`Generated from: ${originalFilename}`, { align: 'center' });
        
        doc.moveDown();
        doc.text(`Date: ${new Date().toLocaleDateString('en-US', { 
          weekday: 'long', 
          year: 'numeric', 
          month: 'long', 
          day: 'numeric' 
        })}`, { align: 'center' });
        
        // Add each graph on a new page with improved layout
        graphs.forEach((graph, index) => {
          doc.addPage();
          
          // Add page header with better styling
          doc.fontSize(14)
             .font('Helvetica-Bold')
             .fillColor('#374151')
             .text(`Graph ${index + 1} of ${graphs.length}`, 50, 40);
          
          // Add a subtle line under header
          doc.moveTo(50, 65)
             .lineTo(doc.page.width - 50, 65)
             .strokeColor('#e5e7eb')
             .lineWidth(1)
             .stroke();
          
          // Check if image exists
          if (fs.existsSync(graph.imagePath)) {
            try {
              // Calculate image dimensions to fit on page while maintaining quality
              const pageWidth = doc.page.width - 100; // 50 margin on each side
              const pageHeight = doc.page.height - 200; // Space for header and footer
              
              // Add the graph image with high quality settings
              doc.image(graph.imagePath, 50, 85, {
                fit: [pageWidth, pageHeight],
                align: 'center',
                valign: 'center',
                quality: 100  // Maximum quality
              });
              
              // Add page footer with better styling
              const footerY = doc.page.height - 60;
              
              // Add subtle line above footer
              doc.moveTo(50, footerY - 10)
                 .lineTo(doc.page.width - 50, footerY - 10)
                 .strokeColor('#e5e7eb')
                 .lineWidth(1)
                 .stroke();
              
              // Page number
              doc.fontSize(10)
                 .font('Helvetica')
                 .fillColor('#9ca3af')
                 .text(`Page ${index + 2}`, 50, footerY, {
                   align: 'center',
                   width: pageWidth
                 });
                 
            } catch (imageError) {
              console.error(`Error adding image ${graph.imagePath}:`, imageError);
              doc.fontSize(14)
                 .fillColor('#ef4444')
                 .text(`Error loading graph: ${graph.title}`, 50, 300, {
                   align: 'center'
                 });
            }
          } else {
            doc.fontSize(14)
               .fillColor('#ef4444')
               .text(`Graph not found: ${graph.title}`, 50, 300, {
                 align: 'center'
               });
          }
        });

        // Add final page with summary
        doc.addPage();
        doc.fontSize(20)
           .font('Helvetica-Bold')
           .fillColor('#1f2937')
           .text('Summary', { align: 'center' });
           
        doc.moveDown();
        doc.fontSize(12)
           .font('Helvetica')
           .fillColor('#4b5563')
           .text(`Total graphs generated: ${graphs.length}`, { align: 'center' });
           
        doc.moveDown();
        doc.text('This document contains statistical analysis graphs generated from the provided data tables.', {
          align: 'center',
          width: doc.page.width - 100
        });

        // Finalize the PDF
        doc.end();

        stream.on('finish', () => {
          resolve(outputPath);
        });

        stream.on('error', (error) => {
          reject(error);
        });

      } catch (error) {
        reject(error);
      }
    });
  }
}

module.exports = new PDFService();