const mammoth = require('mammoth');
const fs = require('fs').promises;

class SignsParserService {
  async parseSignsDocument(filePath) {
    try {
      const buffer = await fs.readFile(filePath);
      const result = await mammoth.extractRawText({ buffer: buffer });
      const html = await mammoth.convertToHtml({ buffer: buffer });
      
      console.log('Signs document raw text preview:', result.value.substring(0, 500));
      
      // Try to extract tables from HTML first
      const tables = this.extractTablesFromHtml(html.value);
      
      if (tables.length === 0) {
        // Fallback: try to parse from raw text
        console.log('HTML parsing failed, trying text parsing...');
        return this.extractTablesFromText(result.value);
      }
      
      return tables;
    } catch (error) {
      console.error('Error parsing Signs document:', error);
      throw new Error('Failed to parse Signs document: ' + error.message);
    }
  }

  extractTablesFromHtml(html) {
    const tables = [];
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    const matches = html.matchAll(tableRegex);

    console.log('HTML content preview:', html.substring(0, 1000));

    for (const match of matches) {
      const tableHtml = match[0];
      console.log('Processing HTML table:', tableHtml.substring(0, 300));
      const tableData = this.parseHtmlSignsTable(tableHtml);
      if (tableData) {
        console.log('Successfully parsed table:', tableData.title);
        tables.push(tableData);
      }
    }

    console.log(`Found ${tables.length} tables from HTML`);
    return tables;
  }

  parseHtmlSignsTable(tableHtml) {
    const rows = [];
    const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    
    const rowMatches = tableHtml.matchAll(rowRegex);
    
    for (const rowMatch of rowMatches) {
      const cells = [];
      const cellMatches = rowMatch[1].matchAll(cellRegex);
      
      for (const cellMatch of cellMatches) {
        // Remove HTML tags and clean text
        const text = cellMatch[1]
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/\^th\^|\^st\^|\^nd\^|\^rd\^/g, '') // Remove superscript markers
          .trim();
        cells.push(text);
      }
      
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    console.log('HTML table rows:', rows);
    return this.processSignsTableRows(rows);
  }

  extractTablesFromText(text) {
    const tables = [];
    const lines = text.split('\n').map(line => line.trim()).filter(line => line);
    
    console.log('Text parsing - total lines:', lines.length);
    console.log('First 10 lines:', lines.slice(0, 10));
    
    let currentTableLines = [];
    let tableNumber = null;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      // Detect table start - more flexible patterns
      if ((line.toLowerCase().includes('table') && line.includes(':')) ||
          line.toLowerCase().includes('signs') && line.toLowerCase().includes('symptoms')) {
        
        // Process previous table if exists
        if (currentTableLines.length > 0) {
          const tableData = this.parseTextTable(currentTableLines, tableNumber);
          if (tableData) {
            console.log('Parsed table from text:', tableData.title);
            tables.push(tableData);
          }
        }
        
        // Extract table number
        const numberMatch = line.match(/table\s*(?:number)?\s*:?\s*(\d+)/i);
        tableNumber = numberMatch ? parseInt(numberMatch[1]) : tables.length + 1;
        currentTableLines = [];
        
        console.log('Found table start:', line, 'Table number:', tableNumber);
      } else if (line.includes('|') || line.includes('+') || line.includes('-')) {
        // This looks like table content
        currentTableLines.push(line);
      } else if (currentTableLines.length > 0 && 
                (line.toLowerCase().includes('vedana') || 
                 line.toLowerCase().includes('varna') ||
                 line.toLowerCase().includes('assessment') ||
                 /\d+(\.\d+)?%/.test(line))) {
        // This might be table data without pipes
        currentTableLines.push(line);
      }
    }
    
    // Process the last table
    if (currentTableLines.length > 0) {
      const tableData = this.parseTextTable(currentTableLines, tableNumber);
      if (tableData) {
        console.log('Parsed final table from text:', tableData.title);
        tables.push(tableData);
      }
    }
    
    console.log(`Found ${tables.length} tables from text`);
    return tables;
  }

  parseTextTable(lines, tableNumber) {
    try {
      console.log('Parsing text table:', lines);
      
      // Extract data rows (skip header and separator lines)
      const dataLines = lines.filter(line => 
        !line.toLowerCase().includes('signs & symptoms') &&
        !line.toLowerCase().includes('assessment') &&
        !line.includes('===') &&
        !line.includes('+++') &&
        (line.includes('|') || /\d+(\.\d+)?%/.test(line)) &&
        !line.match(/^\s*\+[-=+]*\+\s*$/)
      );

      console.log('Data lines:', dataLines);

      if (dataLines.length === 0) return null;

      const signsData = {};
      let currentSign = null;

      for (const line of dataLines) {
        if (line.includes('|')) {
          // Pipe-separated format
          const cells = line.split('|').map(cell => cell.trim()).filter(cell => cell);
          console.log('Processing cells:', cells);
          
          if (cells.length >= 3) {
            // Check if this is a new sign (first column has content)
            if (cells[0] && cells[0] !== '') {
              currentSign = cells[0];
              signsData[currentSign] = {};
            }
            
            // Extract assessment and percentage
            if (currentSign && cells[1] && cells[2]) {
              const assessment = this.normalizeAssessment(cells[1]);
              const percentage = this.extractPercentage(cells[2]);
              
              if (percentage !== null) {
                signsData[currentSign][assessment] = percentage;
                console.log(`Added: ${currentSign}[${assessment}] = ${percentage}%`);
              }
            }
          }
        } else {
          // Try to parse non-pipe format
          const percentMatch = line.match(/(\d+(?:\.\d+)?)%/);
          const dayMatch = line.match(/(\d+)(?:th|st|nd|rd)?\s*day/i);
          
          if (percentMatch && dayMatch) {
            const percentage = parseFloat(percentMatch[1]);
            const assessment = `${dayMatch[1]} day`;
            
            // Try to find the sign name in the line
            const signMatch = line.match(/(vedana|varna|sraava|gandha|maamasankura|parimana)/i);
            if (signMatch) {
              const sign = signMatch[1];
              if (!signsData[sign]) {
                signsData[sign] = {};
              }
              signsData[sign][assessment] = percentage;
              console.log(`Added (non-pipe): ${sign}[${assessment}] = ${percentage}%`);
            }
          }
        }
      }

      console.log('Final signsData:', signsData);

      // Convert to required format
      return this.formatSignsData(signsData, tableNumber);
      
    } catch (error) {
      console.error('Error parsing text table:', error);
      return null;
    }
  }

  processSignsTableRows(rows) {
    if (rows.length < 2) return null;

    console.log('Processing signs table rows:', rows);

    // Find header structure
    let headerRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (row.some(cell => cell.toLowerCase().includes('signs') && cell.toLowerCase().includes('symptoms'))) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      console.log('No header row found, assuming first row is header');
      headerRowIndex = 0;
    }

    // Process data rows
    const signsData = {};
    let currentSign = null;

    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      console.log('Processing row:', row);
      
      // Determine if this is a new sign row or continuation row
      if (row.length >= 3 && row[0] && row[0].trim() !== '') {
        // This is a new sign (3-column format with sign name)
        currentSign = row[0].trim();
        if (!signsData[currentSign]) {
          signsData[currentSign] = {};
        }
        
        // Process the data in this row
        const assessment = this.normalizeAssessment(row[1]);
        const percentage = this.extractPercentage(row[2]);
        
        if (percentage !== null) {
          signsData[currentSign][assessment] = percentage;
          console.log(`Added: ${currentSign}[${assessment}] = ${percentage}%`);
        }
        
      } else if (row.length === 2 && currentSign) {
        // This is a continuation row (2-column format: Assessment, Percentage)
        const assessment = this.normalizeAssessment(row[0]);
        const percentage = this.extractPercentage(row[1]);
        
        if (percentage !== null) {
          signsData[currentSign][assessment] = percentage;
          console.log(`Added: ${currentSign}[${assessment}] = ${percentage}%`);
        }
        
      } else if (row.length >= 3 && (!row[0] || row[0].trim() === '') && currentSign) {
        // This is a continuation row (3-column format with empty first column)
        const assessment = this.normalizeAssessment(row[1]);
        const percentage = this.extractPercentage(row[2]);
        
        if (percentage !== null) {
          signsData[currentSign][assessment] = percentage;
          console.log(`Added: ${currentSign}[${assessment}] = ${percentage}%`);
        }
      }
    }

    console.log('Final processed signsData:', signsData);
    return this.formatSignsData(signsData, 1);
  }

  normalizeAssessment(assessmentText) {
    // Remove ordinal suffixes and normalize
    return assessmentText
      .replace(/th|st|nd|rd/g, '')
      .replace(/(\d+)\s*day/i, '$1 day')
      .trim();
  }

  extractPercentage(percentText) {
    // Extract percentage value, handling various formats
    const cleaned = percentText.replace(/[%>\s]/g, '').trim();
    const percentage = parseFloat(cleaned);
    
    if (isNaN(percentage)) {
      console.log('Could not parse percentage from:', percentText);
      return null;
    }
    
    return percentage;
  }

  formatSignsData(signsData, tableNumber = 1) {
    const signs = Object.keys(signsData);
    
    if (signs.length === 0) {
      console.log('No signs data found');
      return null;
    }

    // Get all unique assessments across all signs
    const allAssessments = new Set();
    Object.values(signsData).forEach(assessments => {
      Object.keys(assessments).forEach(assessment => allAssessments.add(assessment));
    });

    const assessmentOrder = Array.from(allAssessments).sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)?.[0] || '0');
      const numB = parseInt(b.match(/\d+/)?.[0] || '0');
      return numA - numB;
    });

    console.log('Formatted signs:', signs);
    console.log('Assessment order:', assessmentOrder);
    console.log('Data structure:', signsData);

    return {
      title: `Signs & Symptoms Assessment - Table ${tableNumber}`,
      tableNumber: tableNumber,
      signs: signs,
      assessments: assessmentOrder,
      data: signsData,
      type: 'signs_symptoms'
    };
  }

  // Validate signs data structure
  validateSignsData(tables) {
    const errors = [];

    if (!tables || tables.length === 0) {
      errors.push('No tables found in the Signs document');
      return { valid: false, errors };
    }

    tables.forEach((table, index) => {
      if (!table.signs || table.signs.length === 0) {
        errors.push(`Table ${index + 1}: No signs/symptoms found`);
      }

      if (!table.assessments || table.assessments.length === 0) {
        errors.push(`Table ${index + 1}: No assessment periods found`);
      }

      if (!table.data || Object.keys(table.data).length === 0) {
        errors.push(`Table ${index + 1}: No percentage data found`);
      }

      // Check data completeness
      if (table.signs && table.assessments) {
        table.signs.forEach(sign => {
          const signData = table.data[sign] || {};
          table.assessments.forEach(assessment => {
            if (!(assessment in signData)) {
              errors.push(`Table ${index + 1}: Missing data for ${sign} at ${assessment}`);
            }
          });
        });
      }
    });

    return {
      valid: errors.length === 0,
      errors
    };
  }
}

module.exports = new SignsParserService();