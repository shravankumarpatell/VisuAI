const mammoth = require('mammoth');
const fs = require('fs').promises;

class WordParserService {
  async parseWordDocument(filePath) {
    try {
      const buffer = await fs.readFile(filePath);
      const result = await mammoth.extractRawText({ buffer: buffer });
      const html = await mammoth.convertToHtml({ buffer: buffer });
      
      // Try to extract tables from HTML
      const tables = this.extractTablesFromHtml(html.value);
      
      if (tables.length === 0) {
        // Fallback: try to parse from raw text
        return this.extractTablesFromText(result.value);
      }
      
      return tables;
    } catch (error) {
      console.error('Error parsing Word document:', error);
      throw new Error('Failed to parse Word document: ' + error.message);
    }
  }

  // NEW: Parse Word document for improvement percentage data
  async parseImprovementDocument(filePath) {
    try {
      console.log('Parsing improvement document:', filePath);
      
      const buffer = await fs.readFile(filePath);
      const result = await mammoth.extractRawText({ buffer: buffer });
      const html = await mammoth.convertToHtml({ buffer: buffer });
      
      // Try HTML parsing first
      let improvementData = this.extractImprovementTableFromHtml(html.value);
      
      if (!improvementData) {
        // Fallback: try text parsing
        improvementData = this.extractImprovementTableFromText(result.value);
      }
      
      if (!improvementData) {
        throw new Error('No valid improvement percentage table found in document. Expected table with improvement categories and Group A/B data for different time periods.');
      }
      
      console.log('Successfully parsed improvement data:', {
        categories: improvementData.categories.length,
        timePeriods: improvementData.timePeriods.length,
        hasGroupA: !!improvementData.groupA,
        hasGroupB: !!improvementData.groupB
      });
      
      return improvementData;
      
    } catch (error) {
      console.error('Error parsing improvement document:', error);
      throw new Error('Failed to parse improvement document: ' + error.message);
    }
  }

  extractImprovementTableFromHtml(html) {
    try {
      // Find table with improvement data
      const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
      const matches = html.matchAll(tableRegex);

      for (const match of matches) {
        const tableHtml = match[0];
        const improvementData = this.parseImprovementHtmlTable(tableHtml);
        if (improvementData) {
          return improvementData;
        }
      }
      
      return null;
    } catch (error) {
      console.error('Error extracting improvement table from HTML:', error);
      return null;
    }
  }

  parseImprovementHtmlTable(tableHtml) {
    const rows = [];
    const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    
    const rowMatches = tableHtml.matchAll(rowRegex);
    
    for (const rowMatch of rowMatches) {
      const cells = [];
      const cellMatches = rowMatch[1].matchAll(cellRegex);
      
      for (const cellMatch of cellMatches) {
        const text = cellMatch[1]
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .trim();
        cells.push(text);
      }
      
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    if (rows.length < 3) return null; // Need at least header + 2 data rows

    return this.processImprovementTableRows(rows);
  }

  extractImprovementTableFromText(text) {
    try {
      const lines = text.split('\n').map(line => line.trim()).filter(line => line);
      
      // Look for improvement table indicators
      let tableStartIndex = -1;
      let tableEndIndex = -1;
      
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].toLowerCase();
        
        // Look for table with improvement categories
        if (line.includes('improvement') || 
            (line.includes('cured') && line.includes('group')) ||
            (line.includes('marked') && line.includes('improved'))) {
          tableStartIndex = i;
          break;
        }
      }
      
      if (tableStartIndex === -1) return null;
      
      // Find table end (look for next table or significant gap)
      for (let i = tableStartIndex + 1; i < lines.length; i++) {
        const line = lines[i].toLowerCase();
        if (line.includes('table no') || line.includes('graph no') || line === '') {
          tableEndIndex = i;
          break;
        }
      }
      
      if (tableEndIndex === -1) tableEndIndex = lines.length;
      
      // Extract table lines
      const tableLines = lines.slice(tableStartIndex, tableEndIndex);
      const tableRows = [];
      
      for (const line of tableLines) {
        // Split by multiple spaces or tabs
        const parts = line.split(/\s{2,}|\t+/).map(p => p.trim()).filter(p => p);
        if (parts.length >= 5) { // Should have category + at least 4 columns of data
          tableRows.push(parts);
        }
      }
      
      if (tableRows.length < 3) return null;
      
      return this.processImprovementTableRows(tableRows);
      
    } catch (error) {
      console.error('Error extracting improvement table from text:', error);
      return null;
    }
  }

  processImprovementTableRows(rows) {
    try {
      if (!rows || rows.length < 2) return null;
      
      const headerRow = rows[0];
      const dataRows = rows.slice(1);
      
      // Identify time periods from header
      const timePeriods = [];
      const groupAColumns = [];
      const groupBColumns = [];
      
      // Parse header to find Group A and Group B columns for each time period
      for (let i = 1; i < headerRow.length; i++) {
        const cell = headerRow[i].toLowerCase();
        
        // Look for day patterns: "7th day", "14th day", etc.
        const dayMatch = cell.match(/\b(7|14|21|28)(?:th)?\s*day\b/);
        if (dayMatch) {
          const day = `${dayMatch[1]}th`;
          if (!timePeriods.includes(day)) {
            timePeriods.push(day);
          }
          
          // Determine if this is Group A or Group B based on position or header text
          if (cell.includes('group a') || cell.includes('groupa') || 
              (timePeriods.length === 1 && groupAColumns.length === 0)) {
            groupAColumns.push({ day: day, index: i });
          } else if (cell.includes('group b') || cell.includes('groupb')) {
            groupBColumns.push({ day: day, index: i });
          } else {
            // Alternate assignment if not explicitly labeled
            if (groupAColumns.filter(col => col.day === day).length === 0) {
              groupAColumns.push({ day: day, index: i });
            } else {
              groupBColumns.push({ day: day, index: i });
            }
          }
        }
      }
      
      // Default time periods if not found in header
      if (timePeriods.length === 0) {
        timePeriods.push('7th', '14th', '21st', '28th');
      }
      
      // Identify improvement categories
      const categories = [];
      const improvementData = {
        categories: [],
        timePeriods: timePeriods,
        groupA: {},
        groupB: {}
      };
      
      // Initialize data structure
      timePeriods.forEach(period => {
        improvementData.groupA[period] = {};
        improvementData.groupB[period] = {};
      });
      
      // Process data rows
      for (const row of dataRows) {
        if (row.length < 2) continue;
        
        const category = row[0].trim();
        
        // Skip total rows or empty categories
        if (category.toLowerCase().includes('total') || !category) continue;
        
        // Normalize category names
        const normalizedCategory = this.normalizeImprovementCategory(category);
        if (!normalizedCategory) continue;
        
        categories.push(normalizedCategory);
        
        // Initialize category data
        timePeriods.forEach(period => {
          improvementData.groupA[period][normalizedCategory] = { count: 0, percentage: 0 };
          improvementData.groupB[period][normalizedCategory] = { count: 0, percentage: 0 };
        });
        
        // Parse data cells
        for (let i = 1; i < row.length && i < headerRow.length; i++) {
          const cellValue = row[i].trim();
          const parsed = this.parseCountPercentage(cellValue);
          
          if (!parsed) continue;
          
          // Determine which time period and group this column represents
          const dayMatch = headerRow[i].match(/\b(7|14|21|28)(?:th)?\s*day\b/);
          if (!dayMatch) continue;
          
          const day = `${dayMatch[1]}th`;
          const isGroupA = headerRow[i].toLowerCase().includes('group a') || 
                          headerRow[i].toLowerCase().includes('groupa') || 
                          (i <= Math.ceil(headerRow.length / 2));
          
          const group = isGroupA ? 'groupA' : 'groupB';
          improvementData[group][day][normalizedCategory] = parsed;
        }
      }
      
      improvementData.categories = [...new Set(categories)];
      
      // Validate we have the expected structure
      if (improvementData.categories.length === 0 || improvementData.timePeriods.length === 0) {
        return null;
      }
      
      return improvementData;
      
    } catch (error) {
      console.error('Error processing improvement table rows:', error);
      return null;
    }
  }

  normalizeImprovementCategory(category) {
    const cat = category.toLowerCase().trim();
    
    if (cat.includes('cured') && cat.includes('100')) {
      return 'Cured(100%)';
    } else if (cat.includes('marked') && (cat.includes('75') || cat.includes('improve'))) {
      return 'Marked improved(75-100%)';
    } else if (cat.includes('moderate') && (cat.includes('50') || cat.includes('improve'))) {
      return 'Moderate improved(50-75%)';
    } else if (cat.includes('mild') && (cat.includes('25') || cat.includes('improve'))) {
      return 'Mild improved(25-50%)';
    } else if (cat.includes('not cured') || (cat.includes('25') && cat.includes('<'))) {
      return 'Not cured(<25%)';
    }
    
    // Return original if no match (for flexibility)
    return category.trim();
  }

  parseCountPercentage(cellValue) {
    if (!cellValue) return null;
    
    // Handle formats like "8 (40%)", "5(25%)", "0 (0%)", etc.
    const match = cellValue.match(/(\d+)\s*\((\d+(?:\.\d+)?)\s*%?\)/);
    if (match) {
      return {
        count: parseInt(match[1], 10),
        percentage: parseFloat(match[2])
      };
    }
    
    // Handle simple numbers
    const numMatch = cellValue.match(/^\d+$/);
    if (numMatch) {
      return {
        count: parseInt(cellValue, 10),
        percentage: 0
      };
    }
    
    return null;
  }

  // EXISTING METHODS (unchanged)
  extractTablesFromHtml(html) {
    const tables = [];
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    const matches = html.matchAll(tableRegex);

    for (const match of matches) {
      const tableHtml = match[0];
      const tableData = this.parseHtmlTable(tableHtml);
      if (tableData) {
        tables.push(tableData);
      }
    }

    return tables;
  }

  parseHtmlTable(tableHtml) {
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
          .trim();
        cells.push(text);
      }
      
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    if (rows.length < 2) return null;

    // Extract title from the table or use first column header
    let title = 'Table';
    const headerRow = rows[0];
    
    // Check if this looks like our expected table format
    if (headerRow.length >= 5 && 
        headerRow[1].toLowerCase().includes('trial') && 
        headerRow[2].toLowerCase().includes('control')) {
      
      title = headerRow[0] || 'Data';
      
      const data = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row.length >= 5) {
          data.push({
            category: row[0],
            trialGroup: parseInt(row[1]) || 0,
            controlGroup: parseInt(row[2]) || 0,
            total: parseInt(row[3]) || 0,
            percentage: parseFloat(row[4].replace('%', '')) || 0
          });
        }
      }
      
      return {
        title: title,
        rows: data
      };
    }

    return null;
  }

  extractTablesFromText(text) {
    const tables = [];
    const lines = text.split('\n').map(line => line.trim()).filter(line => line);
    
    let currentTable = null;
    let isInTable = false;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      // Look for table headers
      if (line.toLowerCase().includes('table no') && line.includes(':')) {
        if (currentTable && currentTable.rows.length > 0) {
          tables.push(currentTable);
        }
        
        const titleMatch = line.match(/:\s*(.+)/);
        const title = titleMatch ? titleMatch[1].replace(/[:\s]+$/, '') : 'Table';
        
        currentTable = {
          title: title,
          rows: []
        };
        isInTable = false;
      }
      
      // Look for table data start
      if (currentTable && !isInTable && 
          line.toLowerCase().includes('trial group') && 
          line.toLowerCase().includes('control group')) {
        isInTable = true;
        continue;
      }
      
      // Parse table rows
      if (currentTable && isInTable) {
        const parts = line.split(/\s{2,}|\t+/);
        
        if (parts.length >= 5) {
          const row = {
            category: parts[0],
            trialGroup: parseInt(parts[1]) || 0,
            controlGroup: parseInt(parts[2]) || 0,
            total: parseInt(parts[3]) || 0,
            percentage: parseFloat(parts[4].replace('%', '')) || 0
          };
          
          currentTable.rows.push(row);
          
          // Check if this is the total row
          if (parts[0].toLowerCase() === 'total') {
            isInTable = false;
          }
        } else if (line.toLowerCase().includes('graph no') || 
                   line.toLowerCase().includes('table no') ||
                   line === '') {
          isInTable = false;
        }
      }
    }
    
    // Add the last table if exists
    if (currentTable && currentTable.rows.length > 0) {
      tables.push(currentTable);
    }
    
    return tables;
  }
}

module.exports = new WordParserService();