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