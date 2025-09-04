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

  // IMPROVED: Parse Word document for improvement percentage data with dynamic day detection
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
        let text = cellMatch[1]
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '')
          .trim();
        
        // DECODE HTML ENTITIES HERE TOO
        text = this.decodeHtmlEntities(text);
        
        cells.push(text);
      }
      
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    if (rows.length < 3) return null;

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
        // Better splitting logic for text tables
        const parts = line.split(/\s{2,}|\t+/).map(p => p.trim()).filter(p => p);
        
        // More flexible column count check
        if (parts.length >= 3) { // Minimum: category + at least 2 columns of data
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

  // COMPLETELY REWRITTEN: Dynamic day detection with comprehensive pattern matching
// UPDATED: Enhanced processImprovementTableRows method with flexible time period detection
processImprovementTableRows(rows) {
  try {
    if (!rows || rows.length < 2) return null;
    
    const headerRow = rows[0];
    const dataRows = rows.slice(1);
    
    console.log('Processing improvement table with header:', headerRow);
    console.log('Number of data rows:', dataRows.length);
    
    // ENHANCED: More flexible time period detection
    const timePeriods = new Set();
    const groupAColumns = [];
    const groupBColumns = [];
    
    // Parse header to find Group A and Group B columns for each time period
    for (let i = 1; i < headerRow.length; i++) {
      const cell = headerRow[i].toLowerCase().trim();
      
      // ENHANCED: Comprehensive pattern matching for various time formats
      const timePatterns = [
        // Standard day patterns: "7th day", "14th day", "1st day", "3rd day"
        /\b(\d{1,2})(?:st|nd|rd|th)?\s*day\b/i,
        // Reverse patterns: "day 7", "day 14", "day 1", "day 3"
        /\bday\s*(\d{1,2})(?:st|nd|rd|th)?\b/i,
        // Just ordinal numbers: "1st", "3rd", "5th", "7th", "19th"
        /\b(\d{1,2})(?:st|nd|rd|th)\b/i,
        // Simple letter patterns: "AT", "AF", "BT", "BF" (After Treatment, After Follow-up, etc.)
        /\b([A-Z]{1,3})\b/i,
        // Time-based patterns: "week 1", "month 1", "follow-up", "baseline"
        /\b(?:week|month|wk|mo)\s*(\d+)\b/i,
        /\b(follow-?up|baseline|initial|final)\b/i,
        // Plain numbers when context suggests time periods
        /\b(\d{1,2})\b/i
      ];
      
      let timeMatch = null;
      let timePeriod = null;
      
      // Try each pattern
      for (const pattern of timePatterns) {
        timeMatch = cell.match(pattern);
        if (timeMatch) {
          const matchedValue = timeMatch[1];
          
          // Handle different types of matches
          if (/^\d+$/.test(matchedValue)) {
            // Numeric match - validate reasonable range
            const dayNumber = parseInt(matchedValue);
            if (dayNumber >= 1 && dayNumber <= 365) {
              timePeriod = this.formatDayWithOrdinal(dayNumber);
              break;
            }
          } else if (/^[A-Z]{1,3}$/i.test(matchedValue)) {
            // Letter pattern match (AT, AF, etc.)
            timePeriod = matchedValue.toUpperCase();
            break;
          } else if (/^(follow-?up|baseline|initial|final)$/i.test(matchedValue)) {
            // Special time period names
            timePeriod = matchedValue.toLowerCase().replace('-', '');
            break;
          }
        }
      }
      
      if (timePeriod) {
        timePeriods.add(timePeriod);
        
        console.log(`Found time period: ${timePeriod} in column ${i}: "${cell}"`);
        
        // Enhanced group assignment logic
        if (cell.includes('group a') || cell.includes('groupa') || cell.includes('a group') || 
            cell.includes('trial') || cell.includes('treatment') || cell.includes('test')) {
          groupAColumns.push({ day: timePeriod, index: i });
          console.log(`Assigned to Group A: ${timePeriod}`);
        } else if (cell.includes('group b') || cell.includes('groupb') || cell.includes('b group') ||
                   cell.includes('control') || cell.includes('placebo')) {
          groupBColumns.push({ day: timePeriod, index: i });
          console.log(`Assigned to Group B: ${timePeriod}`);
        } else {
          // Smart auto-assignment based on existing groups
          const existingAForThisTime = groupAColumns.filter(col => col.day === timePeriod);
          const existingBForThisTime = groupBColumns.filter(col => col.day === timePeriod);
          
          if (existingAForThisTime.length === 0) {
            groupAColumns.push({ day: timePeriod, index: i });
            console.log(`Auto-assigned to Group A: ${timePeriod} (column ${i})`);
          } else if (existingBForThisTime.length === 0) {
            groupBColumns.push({ day: timePeriod, index: i });
            console.log(`Auto-assigned to Group B: ${timePeriod} (column ${i})`);
          } else {
            // Balanced assignment
            if (groupAColumns.length <= groupBColumns.length) {
              groupAColumns.push({ day: timePeriod, index: i });
              console.log(`Balanced-assigned to Group A: ${timePeriod} (column ${i})`);
            } else {
              groupBColumns.push({ day: timePeriod, index: i });
              console.log(`Balanced-assigned to Group B: ${timePeriod} (column ${i})`);
            }
          }
        }
      }
    }
    
    // Convert Set to Array with intelligent sorting
    const timePeriodsArray = this.sortTimePeriods(Array.from(timePeriods));
    
    console.log('Final time periods found:', timePeriodsArray);
    console.log('Group A columns:', groupAColumns);
    console.log('Group B columns:', groupBColumns);
    
    // FIXED: Handle minimum 1 time period instead of requiring 3
    if (timePeriodsArray.length === 0) {
      console.warn('No time periods found in header row. Attempting fallback detection...');
      
      // Fallback: look for any patterns in header that might indicate time periods
      for (let i = 1; i < headerRow.length; i++) {
        const cell = headerRow[i];
        
        // Try to extract any meaningful identifiers
        const fallbackPatterns = [
          /\b([A-Z]{1,4})\b/g,  // Any capital letters
          /\b(\d+)\b/g,          // Any numbers
          /\b(pre|post|before|after|initial|final)\b/gi  // Common time indicators
        ];
        
        for (const pattern of fallbackPatterns) {
          const matches = cell.matchAll(pattern);
          for (const match of matches) {
            const value = match[1];
            if (value && !timePeriodsArray.includes(value)) {
              timePeriodsArray.push(value);
              // Assign to groups alternately
              if (groupAColumns.length <= groupBColumns.length) {
                groupAColumns.push({ day: value, index: i });
              } else {
                groupBColumns.push({ day: value, index: i });
              }
            }
          }
        }
      }
    }
    
    // FIXED: Allow single time period analysis
    if (timePeriodsArray.length === 0) {
      console.error('Could not detect any time periods from table header');
      return null;
    }
    
    console.log(`Processing table with ${timePeriodsArray.length} time period(s)`);
    
    // Process categories and data
    const categories = [];
    const improvementData = {
      categories: [],
      timePeriods: timePeriodsArray,
      groupA: {},
      groupB: {}
    };
    
    // Initialize data structure for all time periods
    timePeriodsArray.forEach(period => {
      improvementData.groupA[period] = {};
      improvementData.groupB[period] = {};
    });
    
    // Process data rows
    for (const row of dataRows) {
      if (row.length < 2) continue;
      
      const category = row[0].trim();
      
      // Skip total rows or empty categories
      if (category.toLowerCase().includes('total') || !category) continue;
      
      // More flexible category normalization
      const normalizedCategory = this.normalizeImprovementCategory(category);
      if (!normalizedCategory) continue;
      
      categories.push(normalizedCategory);
      
      // Initialize category data for all time periods
      timePeriodsArray.forEach(period => {
        improvementData.groupA[period][normalizedCategory] = { count: 0, percentage: 0 };
        improvementData.groupB[period][normalizedCategory] = { count: 0, percentage: 0 };
      });
      
      // Parse data cells using column mappings
      groupAColumns.forEach(colInfo => {
        if (colInfo.index < row.length) {
          const cellValue = row[colInfo.index].trim();
          const parsed = this.parseCountPercentage(cellValue);
          if (parsed) {
            improvementData.groupA[colInfo.day][normalizedCategory] = parsed;
            console.log(`Group A ${colInfo.day} ${normalizedCategory}: ${parsed.count} (${parsed.percentage}%)`);
          }
        }
      });
      
      groupBColumns.forEach(colInfo => {
        if (colInfo.index < row.length) {
          const cellValue = row[colInfo.index].trim();
          const parsed = this.parseCountPercentage(cellValue);
          if (parsed) {
            improvementData.groupB[colInfo.day][normalizedCategory] = parsed;
            console.log(`Group B ${colInfo.day} ${normalizedCategory}: ${parsed.count} (${parsed.percentage}%)`);
          }
        }
      });
    }
    
    improvementData.categories = [...new Set(categories)];
    
    // Validate we have the expected structure
    if (improvementData.categories.length === 0) {
      console.error('No valid improvement categories found');
      return null;
    }
    
    console.log('Successfully processed improvement data:', {
      categories: improvementData.categories,
      timePeriods: improvementData.timePeriods,
      groupAKeys: Object.keys(improvementData.groupA),
      groupBKeys: Object.keys(improvementData.groupB)
    });
    
    return improvementData;
    
  } catch (error) {
    console.error('Error processing improvement table rows:', error);
    return null;
  }
}

// NEW: Intelligent time period sorting
sortTimePeriods(periods) {
  return periods.sort((a, b) => {
    // Extract numeric values for comparison
    const getNumericValue = (period) => {
      const numMatch = period.match(/\d+/);
      return numMatch ? parseInt(numMatch[0]) : 999;
    };
    
    const getTypeOrder = (period) => {
      // Prioritize different types of time periods
      if (/^\d+(st|nd|rd|th)$/.test(period)) return 1; // Ordinal numbers first
      if (/^[A-Z]{1,3}$/.test(period)) return 2; // Letter codes second
      if (/baseline|initial|pre/i.test(period)) return 0; // Baseline first
      if (/follow|final|post/i.test(period)) return 3; // Follow-up last
      return 2; // Default middle priority
    };
    
    const typeOrderA = getTypeOrder(a);
    const typeOrderB = getTypeOrder(b);
    
    if (typeOrderA !== typeOrderB) {
      return typeOrderA - typeOrderB;
    }
    
    // If same type, sort by numeric value
    return getNumericValue(a) - getNumericValue(b);
  });
}

// ENHANCED: More flexible category normalization
normalizeImprovementCategory(category) {
  const cat = category.toLowerCase().trim();
  
  // Handle various formats of improvement categories
  if (cat.includes('cured') && (cat.includes('100') || cat.includes('complete'))) {
    return 'Cured(100%)';
  } else if (cat.includes('marked') && (cat.includes('75') || cat.includes('improve'))) {
    return 'Marked improved(75-100%)';
  } else if (cat.includes('moderate') && (cat.includes('50') || cat.includes('improve'))) {
    return 'Moderate improved(50-75%)';
  } else if (cat.includes('mild') && (cat.includes('25') || cat.includes('improve'))) {
    return 'Mild improved(25-50%)';
  } else if (cat.includes('not cured') || cat.includes('no improve') || (cat.includes('25') && cat.includes('<'))) {
    return 'Not cured(<25%)';
  } else if (cat.includes('excellent') || cat.includes('complete')) {
    return 'Cured(100%)';
  } else if (cat.includes('good') || cat.includes('significant')) {
    return 'Marked improved(75-100%)';
  } else if (cat.includes('fair') || cat.includes('moderate')) {
    return 'Moderate improved(50-75%)';
  } else if (cat.includes('poor') || cat.includes('minimal')) {
    return 'Mild improved(25-50%)';
  } else if (cat.includes('none') || cat.includes('nil')) {
    return 'Not cured(<25%)';
  }
  
  // Return original if no match (for flexibility)
  return category.trim();
}

  // NEW: Helper method to format day numbers with proper ordinal suffixes
  formatDayWithOrdinal(dayNumber) {
    const num = parseInt(dayNumber);
    
    // Special cases for 11th, 12th, 13th
    if (num % 100 >= 11 && num % 100 <= 13) {
      return `${num}th`;
    }
    
    // Standard ordinal rules
    switch (num % 10) {
      case 1: return `${num}st`;
      case 2: return `${num}nd`;
      case 3: return `${num}rd`;
      default: return `${num}th`;
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
        // Remove HTML tags and decode HTML entities
        let text = cellMatch[1]
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .trim();
        
        // DECODE HTML ENTITIES HERE
        text = this.decodeHtmlEntities(text);
        
        cells.push(text);
      }
      
      if (cells.length > 0) {
        rows.push(cells);
      }
    }

    // Rest of the method remains the same...
    if (rows.length < 2) return null;

    let title = 'Table';
    const headerRow = rows[0];
    
    if (headerRow.length >= 5 && 
        headerRow[1].toLowerCase().includes('trial') && 
        headerRow[2].toLowerCase().includes('control')) {
      
      title = this.decodeHtmlEntities(headerRow[0]) || 'Data';
      
      const data = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (row.length >= 5) {
          data.push({
            category: this.decodeHtmlEntities(row[0]),
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


  decodeHtmlEntities(text) {
    if (!text || typeof text !== 'string') return text;
    
    const entityMap = {
      '&lt;': '<',
      '&gt;': '>',
      '&amp;': '&',
      '&quot;': '"',
      '&#39;': "'",
      '&apos;': "'",
      '&nbsp;': ' ',
      '&copy;': '©',
      '&reg;': '®',
      '&trade;': '™'
    };
    
    return text.replace(/&[a-zA-Z0-9#]+;/g, (entity) => {
      return entityMap[entity] || entity;
    });
  }

  extractTablesFromText(text) {
    // Decode any HTML entities that might exist in extracted text
    text = this.decodeHtmlEntities(text);
    
    const tables = [];
    const lines = text.split('\n').map(line => line.trim()).filter(line => line);
    
    let currentTable = null;
    let isInTable = false;
    
    for (let i = 0; i < lines.length; i++) {
      const line = this.decodeHtmlEntities(lines[i]); // Decode each line
      
      // Look for table headers
      if (line.toLowerCase().includes('table no') && line.includes(':')) {
        if (currentTable && currentTable.rows.length > 0) {
          tables.push(currentTable);
        }
        
        const titleMatch = line.match(/:\s*(.+)/);
        const title = titleMatch ? titleMatch[1].replace(/[:\s]+$/, '') : 'Table';
        
        currentTable = {
          title: this.decodeHtmlEntities(title),
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
            category: this.decodeHtmlEntities(parts[0]),
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