// const { Document, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, TextRun, HeadingLevel, BorderStyle } = require('docx');
// const fs = require('fs').promises;
// const path = require('path');

// class WordService {
//   async generateWordDocument(data, originalFilename) {
//     const doc = new Document({
//       sections: [{
//         properties: {},
//         children: this.createDocumentContent(data)
//       }]
//     });

//     // Generate filename
//     const timestamp = Date.now();
//     const outputPath = path.join('uploads', `tables-${timestamp}.docx`);

//     // Create and save document
//     const buffer = await doc.save();
//     await fs.writeFile(outputPath, buffer);

//     return outputPath;
//   }

//   createDocumentContent(data) {
//     const content = [];

//     // Add title
//     content.push(
//       new Paragraph({
//         text: "Sample Size Estimation Tables",
//         heading: HeadingLevel.HEADING_1,
//         alignment: AlignmentType.CENTER,
//         spacing: {
//           after: 400
//         }
//       })
//     );

//     // Add each table
//     data.tables.forEach((tableData, index) => {
//       // Add table number and title
//       content.push(
//         new Paragraph({
//           text: `Table no ${index + 1}: Distribution of subjects based on ${tableData.title}:`,
//           heading: HeadingLevel.HEADING_2,
//           spacing: {
//             before: 400,
//             after: 200
//           }
//         })
//       );

//       // Create table
//       const table = this.createTable(tableData);
//       content.push(table);

//       // Add spacing after table
//       content.push(
//         new Paragraph({
//           text: "",
//           spacing: {
//             after: 400
//           }
//         })
//       );
//     });

//     return content;
//   }

//   createTable(tableData) {
//     const rows = [];

//     // Header row
//     rows.push(
//       new TableRow({
//         children: [
//           new TableCell({
//             children: [new Paragraph({
//               children: [new TextRun({
//                 text: tableData.title,
//                 bold: true
//               })],
//               alignment: AlignmentType.CENTER
//             })],
//             width: {
//               size: 20,
//               type: WidthType.PERCENTAGE
//             }
//           }),
//           new TableCell({
//             children: [new Paragraph({
//               children: [new TextRun({
//                 text: "Trial Group",
//                 bold: true
//               })],
//               alignment: AlignmentType.CENTER
//             })],
//             width: {
//               size: 20,
//               type: WidthType.PERCENTAGE
//             }
//           }),
//           new TableCell({
//             children: [new Paragraph({
//               children: [new TextRun({
//                 text: "Control Group",
//                 bold: true
//               })],
//               alignment: AlignmentType.CENTER
//             })],
//             width: {
//               size: 20,
//               type: WidthType.PERCENTAGE
//             }
//           }),
//           new TableCell({
//             children: [new Paragraph({
//               children: [new TextRun({
//                 text: "Total",
//                 bold: true
//               })],
//               alignment: AlignmentType.CENTER
//             })],
//             width: {
//               size: 20,
//               type: WidthType.PERCENTAGE
//             }
//           }),
//           new TableCell({
//             children: [new Paragraph({
//               children: [new TextRun({
//                 text: "Percentage (%)",
//                 bold: true
//               })],
//               alignment: AlignmentType.CENTER
//             })],
//             width: {
//               size: 20,
//               type: WidthType.PERCENTAGE
//             }
//           })
//         ]
//       })
//     );

//     // Data rows
//     tableData.rows.forEach((row, index) => {
//       const isLastRow = index === tableData.rows.length - 1;
      
//       rows.push(
//         new TableRow({
//           children: [
//             new TableCell({
//               children: [new Paragraph({
//                 children: [new TextRun({
//                   text: row.category,
//                   bold: isLastRow
//                 })],
//                 alignment: AlignmentType.LEFT
//               })]
//             }),
//             new TableCell({
//               children: [new Paragraph({
//                 children: [new TextRun({
//                   text: row.trialGroup.toString(),
//                   bold: isLastRow
//                 })],
//                 alignment: AlignmentType.CENTER
//               })]
//             }),
//             new TableCell({
//               children: [new Paragraph({
//                 children: [new TextRun({
//                   text: row.controlGroup.toString(),
//                   bold: isLastRow
//                 })],
//                 alignment: AlignmentType.CENTER
//               })]
//             }),
//             new TableCell({
//               children: [new Paragraph({
//                 children: [new TextRun({
//                   text: row.total.toString(),
//                   bold: isLastRow
//                 })],
//                 alignment: AlignmentType.CENTER
//               })]
//             }),
//             new TableCell({
//               children: [new Paragraph({
//                 children: [new TextRun({
//                   text: row.percentage + '%',
//                   bold: isLastRow
//                 })],
//                 alignment: AlignmentType.CENTER
//               })]
//             })
//           ]
//         })
//       );
//     });

//     return new Table({
//       rows: rows,
//       width: {
//         size: 100,
//         type: WidthType.PERCENTAGE
//       },
//       borders: {
//         top: { style: BorderStyle.SINGLE, size: 1 },
//         bottom: { style: BorderStyle.SINGLE, size: 1 },
//         left: { style: BorderStyle.SINGLE, size: 1 },
//         right: { style: BorderStyle.SINGLE, size: 1 },
//         insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
//         insideVertical: { style: BorderStyle.SINGLE, size: 1 }
//       }
//     });
//   }
// }

// module.exports = new WordService();


const { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, TextRun, HeadingLevel } = require('docx');
const fs = require('fs').promises;
const path = require('path');

class WordService {
  async generateWordDocument(data, originalFilename) {
    try {
      console.log('Generating Word document...');
      console.log(`Number of tables to create: ${data.tables ? data.tables.length : 0}`);
      
      const doc = new Document({
        sections: [{
          properties: {},
          children: this.createDocumentContent(data)
        }]
      });

      // Generate filename
      const timestamp = Date.now();
      const outputPath = path.join('uploads', `tables-${timestamp}.docx`);

      // Create buffer using Packer
      console.log('Creating document buffer...');
      const buffer = await Packer.toBuffer(doc);
      
      // Save to file
      console.log('Writing to file...');
      await fs.writeFile(outputPath, buffer);
      
      console.log(`✅ Word document created: ${outputPath}`);
      return outputPath;
      
    } catch (error) {
      console.error('Error in generateWordDocument:', error);
      throw new Error(`Failed to generate Word document: ${error.message}`);
    }
  }

  createDocumentContent(data) {
    const content = [];

    // Add title
    content.push(
      new Paragraph({
        text: "Sample Size Estimation Tables",
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: {
          after: 400
        }
      })
    );

    // Add each table
    if (data.tables && Array.isArray(data.tables)) {
      data.tables.forEach((tableData, index) => {
        // Add table number and title
        content.push(
          new Paragraph({
            text: `Table no ${index + 1}: Distribution of subjects based on ${tableData.title}:`,
            heading: HeadingLevel.HEADING_2,
            spacing: {
              before: 400,
              after: 200
            }
          })
        );

        // Create table
        const table = this.createTable(tableData);
        content.push(table);

        // Add spacing after table
        content.push(
          new Paragraph({
            text: "",
            spacing: {
              after: 400
            }
          })
        );
      });
    } else {
      console.warn('No tables found in data');
      content.push(
        new Paragraph({
          text: "No data available",
          alignment: AlignmentType.CENTER
        })
      );
    }

    return content;
  }

  createTable(tableData) {
    const rows = [];

    // Create header row
    const headerRow = new TableRow({
      children: [
        this.createCell(tableData.title || "Category", true),
        this.createCell("Trial Group", true),
        this.createCell("Control Group", true),
        this.createCell("Total", true),
        this.createCell("Percentage (%)", true)
      ]
    });
    rows.push(headerRow);

    // Create data rows
    if (tableData.rows && Array.isArray(tableData.rows)) {
      tableData.rows.forEach((row, index) => {
        const isLastRow = index === tableData.rows.length - 1;
        
        const dataRow = new TableRow({
          children: [
            this.createCell(String(row.category || ''), isLastRow),
            this.createCell(String(row.trialGroup || 0), isLastRow, true),
            this.createCell(String(row.controlGroup || 0), isLastRow, true),
            this.createCell(String(row.total || 0), isLastRow, true),
            this.createCell(String(row.percentage || '0.00') + '%', isLastRow, true)
          ]
        });
        rows.push(dataRow);
      });
    }

    return new Table({
      rows: rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE
      }
    });
  }

  createCell(text, isBold = false, isCenter = false) {
    return new TableCell({
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: text,
              bold: isBold
            })
          ],
          alignment: isCenter ? AlignmentType.CENTER : AlignmentType.LEFT
        })
      ],
      width: {
        size: 20,
        type: WidthType.PERCENTAGE
      }
    });
  }
}

module.exports = new WordService();