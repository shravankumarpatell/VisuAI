const mammoth = require('mammoth');
const fs = require('fs').promises;

class SignsSymptomsService {
  async parseSignsSymptomsDocument(filePath) {
    try {
      const buffer = await fs.readFile(filePath);
      const result = await mammoth.extractRawText({ buffer: buffer });
      const html = await mammoth.convertToHtml({ buffer: buffer });
      
      // Try to extract tables from HTML first
      let tables = this.extractTablesFromHtml(html.value);
      
      if (tables.length === 0) {
        // Fallback: try to parse from raw text
        tables = this.extractTablesFromText(result.value);
      }
      
      console.log(`Extracted ${tables.length} Signs & Symptoms tables`);
      return tables;
    } catch (error) {
      console.error('Error parsing Signs & Symptoms document:', error);
      throw new Error('Failed to parse Signs & Symptoms document: ' + error.message);
    }
  }

  extractTablesFromHtml(html) {
    const tables = [];
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    const matches = html.matchAll(tableRegex);
    let tableIndex = 1;

    for (const match of matches) {
      const tableHtml = match[0];
      const tableData = this.parseSignsHtmlTable(tableHtml, tableIndex);
      if (tableData && tableData.rows && tableData.rows.length > 0) {
        tables.push(tableData);
        tableIndex++;
      }
    }

    return tables;
  }

  parseSignsHtmlTable(tableHtml, tableIndex) {
    const rows = [];
    const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    
    const rowMatches = tableHtml.matchAll(rowRegex);
    let currentSignSymptom = '';
    let isHeaderRow = true;
    
    for (const rowMatch of rowMatches) {
      const cells = [];
      const cellMatches = rowMatch[1].matchAll(cellRegex);
      
      for (const cellMatch of cellMatches) {
        // Remove HTML tags and clean text
        const text = cellMatch[1]
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/&gt;/g, '>')
          .replace(/\^(st|nd|rd|th)\^/g, '$1') // Remove superscript markers
          .trim();
        cells.push(text);
      }
      
      if (cells.length >= 3 && !isHeaderRow) {
        // Check if this row has a Signs & Symptoms value in first column
        if (cells[0] && cells[0].trim() && !cells[0].toLowerCase().includes('assessment')) {
          currentSignSymptom = cells[0].trim();
        }
        
        // Only process rows with assessment and percentage data
        if (cells[1] && cells[2] && cells[1].toLowerCase().includes('day') && 
            !isNaN(parseFloat(cells[2]))) {
          
          const assessment = cells[1].trim();
          const percentage = parseFloat(cells[2]) || 0;
          
          rows.push({
            signSymptom: currentSignSymptom,
            assessment: assessment,
            percentage: percentage
          });
        }
      }
      
      // Skip header row
      if (isHeaderRow && cells.length >= 3) {
        isHeaderRow = false;
      }
    }

    if (rows.length === 0) return null;

    return {
      title: `Signs and Symptoms Analysis - Group ${tableIndex}`,
      rows: rows
    };
  }

  extractTablesFromText(text) {
    const tables = [];
    const lines = text.split('\n').map(line => line.trim()).filter(line => line);
    
    let currentTable = [];
    let currentSignSymptom = '';
    let tableIndex = 1;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      // Skip header lines
      if (line.toLowerCase().includes('signs') && line.toLowerCase().includes('symptoms') && 
          line.toLowerCase().includes('assessment')) {
        continue;
      }
      
      // Try to parse data lines that contain day assessments and percentages
      const dayMatch = line.match(/(\d+)(?:st|nd|rd|th)?\s*day/i);
      const percentMatch = line.match(/(\d+\.?\d*)\s*%?$/);
      
      if (dayMatch && percentMatch) {
        const assessment = `${dayMatch[1]}th day`;
        const percentage = parseFloat(percentMatch[1]);
        
        if (currentSignSymptom) {
          currentTable.push({
            signSymptom: currentSignSymptom,
            assessment: assessment,
            percentage: percentage
          });
        }
      } else if (line && !line.includes('day') && !line.match(/^\d+\.?\d*\s*%?$/) && 
                 !line.toLowerCase().includes('assessment') && !line.toLowerCase().includes('%')) {
        // This might be a new sign/symptom
        const cleanLine = line.replace(/[>\s]+/g, '').trim();
        if (cleanLine && cleanLine.length > 2) {
          // If we have accumulated data for previous sign/symptom, save it
          if (currentTable.length > 0) {
            // Check if we should start a new table or continue current one
            const uniqueSignsInCurrent = [...new Set(currentTable.map(r => r.signSymptom))];
            if (uniqueSignsInCurrent.length >= 6) { // Threshold for new table
              tables.push({
                title: `Signs and Symptoms Analysis - Group ${tableIndex}`,
                rows: [...currentTable]
              });
              currentTable = [];
              tableIndex++;
            }
          }
          currentSignSymptom = cleanLine;
        }
      }
    }
    
    // Add remaining data as final table
    if (currentTable.length > 0) {
      tables.push({
        title: `Signs and Symptoms Analysis - Group ${tableIndex}`,
        rows: currentTable
      });
    }
    
    return tables;
  }

  // Transform Signs & Symptoms data to graph-compatible format
  transformToGraphFormat(tables) {
    const graphTables = [];
    
    tables.forEach((table, tableIndex) => {
      if (!table.rows || table.rows.length === 0) return;
      
      // Group data by sign/symptom
      const groupedData = {};
      table.rows.forEach(row => {
        if (!groupedData[row.signSymptom]) {
          groupedData[row.signSymptom] = {};
        }
        groupedData[row.signSymptom][row.assessment] = row.percentage;
      });
      
      // Create separate graph for each sign/symptom
      Object.keys(groupedData).forEach(signSymptom => {
        const assessments = groupedData[signSymptom];
        const rows = [];
        
        // Create rows for each assessment day
        Object.keys(assessments).sort((a, b) => {
          const aNum = parseInt(a.match(/\d+/)[0]);
          const bNum = parseInt(b.match(/\d+/)[0]);
          return aNum - bNum;
        }).forEach(assessment => {
          rows.push({
            category: assessment,
            trialGroup: assessments[assessment], // Using percentage as trial group value
            controlGroup: 0, // No control group in this data
            total: assessments[assessment],
            percentage: '100.00' // Each assessment is 100% of itself
          });
        });
        
        // Add total row
        const totalValue = Object.values(assessments).reduce((sum, val) => sum + val, 0);
        rows.push({
          category: 'Total',
          trialGroup: totalValue,
          controlGroup: 0,
          total: totalValue,
          percentage: '100.00'
        });
        
        graphTables.push({
          title: `${signSymptom} - Assessment Timeline`,
          rows: rows,
          colors: {
            trial: '#9fb4ffff',    // Blue for timeline data
            control: '#8e4b6eff',  // Purple (not used but kept for consistency)
            total: '#efe6b2ff'     // Yellow for totals
          },
          xLabel: 'Assessment Timeline',
          yLabel: 'Improvement Percentage (%)'
        });
      });
    });
    
    return graphTables;
  }
}

module.exports = new SignsSymptomsService();