const { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, TextRun, HeadingLevel } = require('docx');
const fs = require('fs').promises;
const path = require('path');

class WordService {
  // Original method for Excel-to-Word tables
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

  // ENHANCED: Generate Master Chart Word Document with statistical tables AND improvement percentage table AND unpaired t-test table
  async generateMasterChartWordDocument(masterChartData, originalFilename) {
    try {
      console.log('Generating Master Chart Word document...');
      console.log(`Number of tables to create: ${masterChartData.tables ? masterChartData.tables.length : 0}`);
      
      // Check if we have improvement analysis data
      const hasImprovementAnalysis = masterChartData.hasImprovementAnalysis || 
        masterChartData.tables.some(table => table.type === 'improvement_percentage');
      
      // Check if we have unpaired t-test data
      const hasUnpairedTTest = masterChartData.hasUnpairedTTest || 
        masterChartData.tables.some(table => table.type === 'unpaired_ttest');
      
      if (hasImprovementAnalysis) {
        console.log('✅ Improvement percentage analysis detected - will generate additional table');
      }
      
      if (hasUnpairedTTest) {
        console.log('✅ Unpaired t-test analysis detected - will generate additional table');
      }
      
      const doc = new Document({
        sections: [{
          properties: {},
          children: this.createMasterChartContent(masterChartData)
        }]
      });

      // Generate filename
      const timestamp = Date.now();
      const outputPath = path.join('uploads', `master-chart-analysis-${timestamp}.docx`);

      // Create buffer using Packer
      console.log('Creating document buffer...');
      const buffer = await Packer.toBuffer(doc);
      
      // Save to file
      console.log('Writing to file...');
      await fs.writeFile(outputPath, buffer);
      
      console.log(`✅ Master Chart Word document created: ${outputPath}`);
      return outputPath;
      
    } catch (error) {
      console.error('Error in generateMasterChartWordDocument:', error);
      throw new Error(`Failed to generate Master Chart Word document: ${error.message}`);
    }
  }

  createMasterChartContent(masterChartData) {
    const content = [];

    // Add main title
    content.push(
      new Paragraph({
        text: "Master Chart Statistical Analysis",
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: {
          after: 600
        }
      })
    );

    // Add source file information
    if (masterChartData.sourceFile) {
      const filename = path.basename(masterChartData.sourceFile);
      content.push(
        new Paragraph({
          children: [
            new TextRun({ text: "Source File: ", bold: true }),
            new TextRun({ text: filename })
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        })
      );
    }

    // Add processing summary
    const statisticalTables = masterChartData.tables.filter(t => t.type === 'statistical_analysis');
    const improvementTables = masterChartData.tables.filter(t => t.type === 'improvement_percentage');
    const unpairedTTestTables = masterChartData.tables.filter(t => t.type === 'unpaired_ttest');
    
    content.push(
      new Paragraph({
        children: [
          new TextRun({ 
            text: `Processed ${masterChartData.processedSheets} of ${masterChartData.totalSheets} sheets • Generated ${statisticalTables.length} statistical tables`,
            italic: true
          }),
          ...(improvementTables.length > 0 ? [
            new TextRun({ text: " • ", italic: true }),
            new TextRun({ text: `${improvementTables.length} improvement percentage analysis table`, italic: true })
          ] : []),
          ...(unpairedTTestTables.length > 0 ? [
            new TextRun({ text: " • ", italic: true }),
            new TextRun({ text: `${unpairedTTestTables.length} unpaired t-test comparison table`, italic: true })
          ] : [])
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 }
      })
    );

    // First, create statistical analysis tables (existing functionality - paired t-test)
    const statisticalTablesBySheet = {};
    statisticalTables.forEach((table, index) => {
      table.tableNumber = (index + 27); // keep starting table numbering at 27 as before
      if (!statisticalTablesBySheet[table.sheetNumber]) {
        statisticalTablesBySheet[table.sheetNumber] = [];
      }
      statisticalTablesBySheet[table.sheetNumber].push(table);
    });

    // Create statistical tables for each sheet
    Object.keys(statisticalTablesBySheet).sort((a, b) => parseInt(a) - parseInt(b)).forEach(sheetNumber => {
      const sheetTables = statisticalTablesBySheet[sheetNumber];
      
      // Add sheet section header
      content.push(
        new Paragraph({
          text: `Sheet ${sheetNumber}: ${sheetTables[0].sheetName}`,
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 600,
            after: 300
          }
        })
      );

      // Add tables for this sheet
      sheetTables.forEach((table) => {
        // Add table title
        content.push(
          new Paragraph({
            children: [
              new TextRun({ 
                text: `Table no-${table.tableNumber}: Comparison of `, 
                bold: true 
              }),
              new TextRun({ 
                text: `${table.title}`, 
                bold: true 
              }),
              new TextRun({ 
                text: ` in signs & symptoms assessment:`, 
                bold: true 
              })
            ],
            spacing: {
              before: 400,
              after: 200
            }
          })
        );

        // Create the statistical table
        const statisticalTable = this.createMasterChartTable(table);
        content.push(statisticalTable);

        // Add spacing after table
        content.push(
          new Paragraph({
            text: "",
            spacing: { after: 400 }
          })
        );
      });
    });

    // NEW: Add unpaired t-test comparison table if available
    const unpairedTable = unpairedTTestTables[0];
    if (unpairedTable) {
      content.push(
        new Paragraph({
          text: "Trial vs Control Group Comparison",
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 800,
            after: 300
          }
        })
      );

      // Add unpaired t-test table title
      content.push(
        new Paragraph({
          children: [
            new TextRun({ 
              text: `Table no-${statisticalTables.length + 27}: ${unpairedTable.title}`, 
              bold: true 
            })
          ],
          spacing: {
            before: 400,
            after: 200
          }
        })
      );

      // Create the unpaired t-test table
      const unpairedTTestTable = this.createUnpairedTTestTable(unpairedTable);
      content.push(unpairedTTestTable);

      // Add spacing after table
      content.push(
        new Paragraph({
          text: "",
          spacing: { after: 400 }
        })
      );
    }

    // Add improvement percentage analysis table if available
    const improvementTable = improvementTables[0];
    if (improvementTable) {
      content.push(
        new Paragraph({
          text: "Patient Improvement Analysis",
          heading: HeadingLevel.HEADING_2,
          spacing: {
            before: 800,
            after: 300
          }
        })
      );

      // Calculate table number (after statistical and unpaired tables)
      const improvementTableNumber = statisticalTables.length + (unpairedTTestTables.length > 0 ? 1 : 0) + 27;

      // Add improvement table title
      content.push(
        new Paragraph({
          children: [
            new TextRun({ 
              text: `Table no-${improvementTableNumber}: ${improvementTable.title}`, 
              bold: true 
            })
          ],
          spacing: {
            before: 400,
            after: 200
          }
        })
      );

      // Create the improvement percentage table
      const improvementPercentageTable = this.createImprovementPercentageTable(improvementTable);
      content.push(improvementPercentageTable);

      // Add spacing after table
      content.push(
        new Paragraph({
          text: "",
          spacing: { after: 400 }
        })
      );
    }

    // Add legend/notes section
    content.push(
      new Paragraph({
        text: "Statistical Notes:",
        heading: HeadingLevel.HEADING_3,
        spacing: {
          before: 600,
          after: 300
        }
      })
    );

    const notes = [
      "B.T = Before Treatment, A.T = After Treatment",
      "df = Degrees of Freedom",
      "Significance Levels: HS = Highly Significant, S = Significant, VS = Very Significant, NS = Not Significant",
      "p < 0.001 = Highly Significant (HS), p < 0.01 = Very Significant (VS), p < 0.05 = Significant (S), p ≥ 0.05 = Not Significant (NS)",
      "Effectiveness % = (Mean Difference / B.T Mean) × 100",
      "Paired t-test analysis performed on before-after treatment data for individual sheets",
    ];

    // Add unpaired t-test notes if we have that table
    if (unpairedTable) {
      notes.push(
        "Unpaired t-test analysis compares Trial Group (Sheet 1) vs Control Group (Sheet 2)",
        "Unpaired t-test df = n1 + n2 - 2, where n1 = trial group size, n2 = control group size"
      );
    }

    // Add improvement percentage notes if we have that table
    if (improvementTable) {
      notes.push(
        "Improvement % = ((BT - AT) / BT) × 100 for each patient",
        "Categories: Cured(100%), Marked improved(75-100%), Moderate improved(50-75%), Mild improved(25-50%), Not cured(<25%)"
      );
    }

    notes.forEach(note => {
      content.push(
        new Paragraph({
          children: [
            new TextRun({ text: "• " }),
            new TextRun({ text: note })
          ],
          spacing: { after: 100 }
        })
      );
    });

    return content;
  }

  // NEW: Create unpaired t-test table
  createUnpairedTTestTable(unpairedTableData) {
    const statsArray = Array.isArray(unpairedTableData.statistics) ? unpairedTableData.statistics : [unpairedTableData.statistics];
    
    const rows = [];

    // Create header row
    const headerRow = new TableRow({
      children: [
        this.createHeaderCell("Symptoms"),
        this.createHeaderCell("Assessment"),
        this.createHeaderCell("Trial\nmean±S.D"),
        this.createHeaderCell("Control\nmean±S.D"),
        this.createHeaderCell("df"),
        this.createHeaderCell("t-value"),
        this.createHeaderCell("p-value"),
        this.createHeaderCell("Remark")
      ]
    });
    rows.push(headerRow);

    // Create data rows for each statistic
    statsArray.forEach((stats, idx) => {
      const assessment = stats.assessment || `Assessment ${idx + 1}`;
      const parameter = stats.parameter || 'Unknown Parameter';

      // Format numeric presentation
      const trialMeanSd = `${this.formatNumber(stats.trialMean, 3)}±${this.formatNumber(stats.trialSD, 4)}`;
      const controlMeanSd = `${this.formatNumber(stats.controlMean, 3)}±${this.formatNumber(stats.controlSD, 4)}`;
      const df = String(stats.degreesOfFreedom);
      const tVal = this.formatNumber(stats.tValue, 4);
      const pVal = this.formatPValue(stats.pValue);
      const remark = this.getSignificanceRemark(stats.pValue);

      const dataRow = new TableRow({
        children: [
          // Symptoms
          this.createDataCell(parameter, false, AlignmentType.LEFT),
          
          // Assessment (7th day / 14th day / etc.)
          this.createDataCell(assessment, false, AlignmentType.CENTER),
          
          // Trial Mean±S.D
          this.createDataCell(trialMeanSd, false, AlignmentType.CENTER),
          
          // Control Mean±S.D
          this.createDataCell(controlMeanSd, false, AlignmentType.CENTER),
          
          // df
          this.createDataCell(df, false, AlignmentType.CENTER),
          
          // t-value
          this.createDataCell(tVal, false, AlignmentType.CENTER),
          
          // p-value
          this.createDataCell(pVal, false, AlignmentType.CENTER),
          
          // Remarks (HS, VS, S, NS)
          this.createDataCell(remark, false, AlignmentType.CENTER)
        ]
      });

      rows.push(dataRow);
    });

    return new Table({
      rows: rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE
      },
      margins: {
        top: 100,
        bottom: 100,
        left: 100,
        right: 100
      }
    });
  }

  // Create improvement percentage table
// Updated createImprovementPercentageTable method in wordService.js

// UPDATED: Create improvement percentage table with support for single group
createImprovementPercentageTable(improvementTableData) {
  const improvementData = improvementTableData.improvementData;
  const timePoints = improvementData.timePoints; // ['7th', '14th', '21st', '28th']
  const categories = improvementData.categories; // ['Cured(100%)', 'Marked improved(75-100%)', etc.]
  const isSingleGroup = improvementData.isSingleGroup || false;
  
  const rows = [];

  // Create header row - adjust based on single or dual group
  const headerCells = [
    this.createHeaderCell("Improvement Category")
  ];
  
  // Add Group A columns
  timePoints.forEach(timePoint => {
    headerCells.push(this.createHeaderCell(`Group A\n${timePoint} day`));
  });
  
  // Only add Group B columns if we have dual group data
  if (!isSingleGroup && improvementData.groupB) {
    timePoints.forEach(timePoint => {
      headerCells.push(this.createHeaderCell(`Group B\n${timePoint} day`));
    });
  }

  const headerRow = new TableRow({
    children: headerCells
  });
  rows.push(headerRow);

  // Create data rows for each improvement category
  categories.forEach(category => {
    const dataCells = [
      this.createDataCell(category, false, AlignmentType.LEFT)
    ];
    
    // Add Group A data
    timePoints.forEach(timePoint => {
      const groupAData = improvementData.groupA[timePoint][category];
      const displayText = `${groupAData.count} (${groupAData.percentage}%)`;
      dataCells.push(this.createDataCell(displayText, false, AlignmentType.CENTER));
    });
    
    // Only add Group B data if we have dual group data
    if (!isSingleGroup && improvementData.groupB) {
      timePoints.forEach(timePoint => {
        const groupBData = improvementData.groupB[timePoint][category];
        const displayText = `${groupBData.count} (${groupBData.percentage}%)`;
        dataCells.push(this.createDataCell(displayText, false, AlignmentType.CENTER));
      });
    }

    const dataRow = new TableRow({
      children: dataCells
    });
    rows.push(dataRow);
  });

  return new Table({
    rows: rows,
    width: {
      size: 100,
      type: WidthType.PERCENTAGE
    },
    margins: {
      top: 100,
      bottom: 100,
      left: 100,
      right: 100
    }
  });
}

// UPDATED: Create Master Chart content with single group support
createMasterChartContent(masterChartData) {
  const content = [];

  // Add main title
  content.push(
    new Paragraph({
      text: "Master Chart Statistical Analysis",
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: {
        after: 600
      }
    })
  );

  // Add source file information
  if (masterChartData.sourceFile) {
    const filename = path.basename(masterChartData.sourceFile);
    content.push(
      new Paragraph({
        children: [
          new TextRun({ text: "Source File: ", bold: true }),
          new TextRun({ text: filename })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 }
      })
    );
  }

  // Add processing summary
  const statisticalTables = masterChartData.tables.filter(t => t.type === 'statistical_analysis');
  const improvementTables = masterChartData.tables.filter(t => t.type === 'improvement_percentage');
  const unpairedTTestTables = masterChartData.tables.filter(t => t.type === 'unpaired_ttest');
  
  content.push(
    new Paragraph({
      children: [
        new TextRun({ 
          text: `Processed ${masterChartData.processedSheets} of ${masterChartData.totalSheets} sheets • Generated ${statisticalTables.length} statistical tables`,
          italic: true
        }),
        ...(improvementTables.length > 0 ? [
          new TextRun({ text: " • ", italic: true }),
          new TextRun({ 
            text: `${improvementTables.length} improvement percentage analysis table${improvementTables[0]?.isSingleGroup ? ' (single group)' : ''}`, 
            italic: true 
          })
        ] : []),
        ...(unpairedTTestTables.length > 0 ? [
          new TextRun({ text: " • ", italic: true }),
          new TextRun({ text: `${unpairedTTestTables.length} unpaired t-test comparison table`, italic: true })
        ] : [])
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 }
    })
  );

  // First, create statistical analysis tables (existing functionality - paired t-test)
  const statisticalTablesBySheet = {};
  statisticalTables.forEach((table, index) => {
    table.tableNumber = (index + 27); // keep starting table numbering at 27 as before
    if (!statisticalTablesBySheet[table.sheetNumber]) {
      statisticalTablesBySheet[table.sheetNumber] = [];
    }
    statisticalTablesBySheet[table.sheetNumber].push(table);
  });

  // Create statistical tables for each sheet
  Object.keys(statisticalTablesBySheet).sort((a, b) => parseInt(a) - parseInt(b)).forEach(sheetNumber => {
    const sheetTables = statisticalTablesBySheet[sheetNumber];
    
    // Add sheet section header
    content.push(
      new Paragraph({
        text: `Sheet ${sheetNumber}: ${sheetTables[0].sheetName}`,
        heading: HeadingLevel.HEADING_2,
        spacing: {
          before: 600,
          after: 300
        }
      })
    );

    // Add tables for this sheet
    sheetTables.forEach((table) => {
      // Add table title
      content.push(
        new Paragraph({
          children: [
            new TextRun({ 
              text: `Table no-${table.tableNumber}: Comparison of `, 
              bold: true 
            }),
            new TextRun({ 
              text: `${table.title}`, 
              bold: true 
            }),
            new TextRun({ 
              text: ` in signs & symptoms assessment:`, 
              bold: true 
            })
          ],
          spacing: {
            before: 400,
            after: 200
          }
        })
      );

      // Create the statistical table
      const statisticalTable = this.createMasterChartTable(table);
      content.push(statisticalTable);

      // Add spacing after table
      content.push(
        new Paragraph({
          text: "",
          spacing: { after: 400 }
        })
      );
    });
  });

  // Add unpaired t-test comparison table if available (only for dual groups)
  const unpairedTable = unpairedTTestTables[0];
  if (unpairedTable) {
    content.push(
      new Paragraph({
        text: "Trial vs Control Group Comparison",
        heading: HeadingLevel.HEADING_2,
        spacing: {
          before: 800,
          after: 300
        }
      })
    );

    // Add unpaired t-test table title
    content.push(
      new Paragraph({
        children: [
          new TextRun({ 
            text: `Table no-${statisticalTables.length + 27}: ${unpairedTable.title}`, 
            bold: true 
          })
        ],
        spacing: {
          before: 400,
          after: 200
        }
      })
    );

    // Create the unpaired t-test table
    const unpairedTTestTable = this.createUnpairedTTestTable(unpairedTable);
    content.push(unpairedTTestTable);

    // Add spacing after table
    content.push(
      new Paragraph({
        text: "",
        spacing: { after: 400 }
      })
    );
  }

  // Add improvement percentage analysis table if available
  const improvementTable = improvementTables[0];
  if (improvementTable) {
    const isSingleGroup = improvementTable.isSingleGroup || false;
    
    content.push(
      new Paragraph({
        text: isSingleGroup ? "Patient Improvement Analysis (Single Group)" : "Patient Improvement Analysis",
        heading: HeadingLevel.HEADING_2,
        spacing: {
          before: 800,
          after: 300
        }
      })
    );

    // Calculate table number (after statistical and unpaired tables)
    const improvementTableNumber = statisticalTables.length + (unpairedTTestTables.length > 0 ? 1 : 0) + 27;

    // Add improvement table title
    content.push(
      new Paragraph({
        children: [
          new TextRun({ 
            text: `Table no-${improvementTableNumber}: ${improvementTable.title}`, 
            bold: true 
          })
        ],
        spacing: {
          before: 400,
          after: 200
        }
      })
    );

    // Create the improvement percentage table
    const improvementPercentageTable = this.createImprovementPercentageTable(improvementTable);
    content.push(improvementPercentageTable);

    // Add spacing after table
    content.push(
      new Paragraph({
        text: "",
        spacing: { after: 400 }
      })
    );
  }

  // Add legend/notes section
  content.push(
    new Paragraph({
      text: "Statistical Notes:",
      heading: HeadingLevel.HEADING_3,
      spacing: {
        before: 600,
        after: 300
      }
    })
  );

  const notes = [
    "B.T = Before Treatment, A.T = After Treatment",
    "df = Degrees of Freedom",
    "Significance Levels: HS = Highly Significant, S = Significant, VS = Very Significant, NS = Not Significant",
    "p < 0.001 = Highly Significant (HS), p < 0.01 = Very Significant (VS), p < 0.05 = Significant (S), p ≥ 0.05 = Not Significant (NS)",
    "Effectiveness % = (Mean Difference / B.T Mean) × 100",
    "Paired t-test analysis performed on before-after treatment data for individual sheets",
  ];

  // Add unpaired t-test notes if we have that table
  if (unpairedTable) {
    notes.push(
      "Unpaired t-test analysis compares Trial Group (Sheet 1) vs Control Group (Sheet 2)",
      "Unpaired t-test df = n1 + n2 - 2, where n1 = trial group size, n2 = control group size"
    );
  }

  // Add improvement percentage notes if we have that table
  if (improvementTable) {
    const isSingleGroup = improvementTable.isSingleGroup || false;
    notes.push(
      "Improvement % = ((BT - AT) / BT) × 100 for each patient",
      "Categories: Cured(100%), Marked improved(75-100%), Moderate improved(50-75%), Mild improved(25-50%), Not cured(<25%)"
    );
    
    if (isSingleGroup) {
      notes.push("Single group analysis: Only Group A data available for improvement categorization");
    }
  }

  notes.forEach(note => {
    content.push(
      new Paragraph({
        children: [
          new TextRun({ text: "• " }),
          new TextRun({ text: note })
        ],
        spacing: { after: 100 }
      })
    );
  });

  return content;
}

  // This function now expects tableData.statistics to be an array (one object per AT subcolumn)
  createMasterChartTable(tableData) {
    const statsArray = Array.isArray(tableData.statistics) ? tableData.statistics : [tableData.statistics];
    
    // Extract parameter name from statistics array or use fallback
    let paramName = 'Unknown Parameter';
    if (statsArray.length > 0 && statsArray[0].parameter) {
      paramName = statsArray[0].parameter;
    } else if (tableData.parameterGroup && tableData.parameterGroup.name) {
      paramName = tableData.parameterGroup.name;
    } else if (tableData.title && !tableData.title.includes('Master Chart')) {
      paramName = tableData.title;
    }
    
    const rows = [];

    // Create header row
    const headerRow = new TableRow({
      children: [
        this.createHeaderCell("Signs & Symptoms"),
        this.createHeaderCell("Assessment"),
        this.createHeaderCell("B.T Mean±S.D"),
        this.createHeaderCell("A.T Mean±S.D"),
        this.createHeaderCell("df"),
        this.createHeaderCell("t-value"),
        this.createHeaderCell("p-value"),
        this.createHeaderCell("Effectiveness %"),
        this.createHeaderCell("Remarks")
      ]
    });
    rows.push(headerRow);

    // For each stats entry (one per AT subcolumn), create a row showing its values.
    statsArray.forEach((stats, idx) => {
      const assessment = stats.assessment || stats.atHeader || `Assessment ${idx + 1}`;

      // Get the parameter name for this specific stat entry
      const currentParamName = stats.parameter || paramName;

      // Format numeric presentation
      const btMeanSd = `${this.formatNumber(stats.meanBT, 3)}±${this.formatNumber(stats.sdBT, 4)}`;
      const atMeanSd = `${this.formatNumber(stats.meanAT, 3)}±${this.formatNumber(stats.sdAT, 4)}`;
      const df = String(stats.degreesOfFreedom);
      const tVal = this.formatNumber(stats.tValue, 4);
      const pVal = this.formatPValue(stats.pValue);
      const eff = `${this.formatNumber(Math.abs(stats.effectiveness), 2)}%`;
      const remark = this.getSignificanceRemark(stats.pValue);

      const dataRow = new TableRow({
        children: [
          // Signs & Symptoms (show parameter name in every row)
          this.createDataCell(currentParamName, false, AlignmentType.LEFT),
          
          // Assessment (7th day / 14th day / etc.)
          this.createDataCell(assessment, false, AlignmentType.CENTER),
          
          // B.T Mean±S.D
          this.createDataCell(btMeanSd, false, AlignmentType.CENTER),
          
          // A.T Mean±S.D
          this.createDataCell(atMeanSd, false, AlignmentType.CENTER),
          
          // df
          this.createDataCell(df, false, AlignmentType.CENTER),
          
          // t-value
          this.createDataCell(tVal, false, AlignmentType.CENTER),
          
          // p-value (without asterisks)
          this.createDataCell(pVal, false, AlignmentType.CENTER),
          
          // Effectiveness %
          this.createDataCell(eff, false, AlignmentType.CENTER),
          
          // Remarks (HS, VS, S, NS)
          this.createDataCell(remark, false, AlignmentType.CENTER)
        ]
      });

      rows.push(dataRow);
    });

    return new Table({
      rows: rows,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE
      },
      margins: {
        top: 100,
        bottom: 100,
        left: 100,
        right: 100
      }
    });
  }

  createHeaderCell(text) {
    return new TableCell({
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: text,
              bold: true,
              size: 20
            })
          ],
          alignment: AlignmentType.CENTER
        })
      ],
      width: {
        size: text === "Signs & Symptoms" ? 18 : 
              text === "Symptoms" ? 18 :
              text === "Assessment" ? 12 : 
              text === "Remarks" ? 10 : 
              text === "Remark" ? 10 :
              text === "Improvement Category" ? 20 : 11,
        type: WidthType.PERCENTAGE
      },
      shading: {
        fill: "E6E6FA"
      }
    });
  }

  createDataCell(text, isBold = false, alignment = AlignmentType.CENTER) {
    return new TableCell({
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: text,
              bold: isBold,
              size: 18
            })
          ],
          alignment: alignment
        })
      ]
    });
  }

  formatNumber(value, decimals = 3) {
    if (typeof value !== 'number' || !isFinite(value)) {
      // preserve same output style as before
      return (decimals === 0) ? '0' : Number(0).toFixed(decimals);
    }
    return value.toFixed(decimals);
  }

  formatPValue(pValue) {
    if (typeof pValue !== 'number' || !isFinite(pValue)) {
      return 'N/A';
    }
    
    if (pValue < 0.001) {
      return '< 0.001';
    } else {
      return pValue.toFixed(3);
    }
  }

  getSignificanceRemark(pValue) {
    if (typeof pValue !== 'number' || !isFinite(pValue)) {
      return 'N/A';
    }
    
    if (pValue < 0.001) {
      return 'HS';  // Highly Significant
    } else if (pValue < 0.01) {
      return 'VS';  // Very Significant
    } else if (pValue < 0.05) {
      return 'S';   // Significant
    } else {
      return 'NS';  // Not Significant
    }
  }

  // Original methods for existing functionality (left intact)
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