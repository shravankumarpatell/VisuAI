// backend/services/masterChartService.js
const ExcelJS = require('exceljs');

class MasterChartService {
  async processMasterChart(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const processedTables = [];
    const groupData = {}; // Store data for improvement percentage calculation

    for (let sheetIndex = 0; sheetIndex < workbook.worksheets.length; sheetIndex++) {
      const worksheet = workbook.worksheets[sheetIndex];
      if (!worksheet || worksheet.actualRowCount === 0) continue;

      const sheetData = await this._extractSheetDataWithMultiRowHeaders(worksheet);
      if (!this._hasValidData(sheetData)) continue;

      // Store sheet data for improvement percentage calculation
      const groupKey = sheetIndex === 0 ? 'groupA' : sheetIndex === 1 ? 'groupB' : `group${String.fromCharCode(65 + sheetIndex)}`;
      groupData[groupKey] = {
        sheetData,
        sheetNumber: sheetIndex + 1,
        sheetName: worksheet.name
      };

      // Generate existing statistical tables (paired t-test)
      const table = await this._processSheetIntoSingleTable(sheetData, sheetIndex + 1, worksheet.name);
      if (table) processedTables.push(table);
    }

    // UPDATED: Generate improvement percentage table if we have at least 1 group (changed from 2)
    if (Object.keys(groupData).length >= 1) {
      const improvementTable = await this._generateImprovementPercentageTable(groupData);
      if (improvementTable) processedTables.push(improvementTable);
    }

    // Generate unpaired t-test table if we have exactly 2 groups (Trial vs Control)
    if (Object.keys(groupData).length >= 2) {
      const unpairedTable = await this._generateUnpairedTTestTable(groupData, processedTables);
      if (unpairedTable) processedTables.push(unpairedTable);
    }

    if (processedTables.length === 0) throw new Error('No valid data found in master chart.');

    return {
      tables: processedTables,
      sourceFile: filePath,
      totalSheets: workbook.worksheets.length,
      processedSheets: Object.keys(groupData).length,
      hasImprovementAnalysis: Object.keys(groupData).length >= 1, // UPDATED: changed from 2 to 1
      hasUnpairedTTest: Object.keys(groupData).length >= 2
    };
  }

  // UPDATED: Generate improvement percentage table for single sheet or multiple sheets
  async _generateImprovementPercentageTable(groupData) {
    try {
      console.log('Generating improvement percentage table...');
      
      // Get available groups
      const groupKeys = Object.keys(groupData);
      const groupA = groupData.groupA;
      const groupB = groupData.groupB || null; // groupB is optional now
      
      if (!groupA) {
        console.warn('Need at least 1 sheet (Group A) for improvement percentage analysis');
        return null;
      }

      // Identify parameters in Group A
      const groupAParams = this._identifyParameterGroups(groupA.sheetData.headers);
      let groupBParams = [];
      let commonParams = [];

      if (groupB) {
        // If we have Group B, find common parameters
        groupBParams = this._identifyParameterGroups(groupB.sheetData.headers);
        commonParams = this._findCommonParameters(groupAParams, groupBParams);
        
        if (commonParams.length === 0) {
          console.warn('No common parameters found between Group A and Group B, using Group A parameters only');
          commonParams = groupAParams.map(param => ({
            name: param.name,
            groupA: param,
            groupB: null
          }));
        }
      } else {
        // Single sheet mode - use all Group A parameters
        console.log('Single sheet mode: Using Group A parameters only');
        commonParams = groupAParams.map(param => ({
          name: param.name,
          groupA: param,
          groupB: null
        }));
      }

      if (commonParams.length === 0) {
        console.warn('No parameters found for improvement percentage analysis');
        return null;
      }

      console.log(`Found ${commonParams.length} parameters for improvement analysis`);

      // Calculate improvement percentages
      const improvementData = this._calculateImprovementPercentages(
        commonParams, 
        groupA.sheetData, 
        groupB ? groupB.sheetData : null
      );

      // Create the improvement percentage table
      const improvementTable = {
        title: groupB ? 
          'Patient Improvement Analysis by Categories' : 
          'Patient Improvement Analysis by Categories (Single Group)',
        type: 'improvement_percentage',
        sheetNumber: 999, // Special number to distinguish from regular tables
        sheetName: 'Improvement Analysis',
        improvementData: improvementData,
        tableNumber: 999,
        isSingleGroup: !groupB // Flag to indicate single group analysis
      };

      return improvementTable;

    } catch (error) {
      console.error('Error generating improvement percentage table:', error);
      return null;
    }
  }

  // UPDATED: Calculate improvement percentages with support for single group
  _calculateImprovementPercentages(commonParams, groupAData, groupBData = null) {
    const timePoints = ['7th', '14th', '21st', '28th'];
    const categories = [
      { name: 'Cured(100%)', min: 100, max: Infinity },
      { name: 'Marked improved(75-100%)', min: 75, max: 100 },
      { name: 'Moderate improved(50-75%)', min: 50, max: 75 },
      { name: 'Mild improved(25-50%)', min: 25, max: 50 },
      { name: 'Not cured(<25%)', min: -Infinity, max: 25 }
    ];

    const result = {
      groupA: {},
      groupB: groupBData ? {} : null, // Only create groupB if we have data
      timePoints,
      categories: categories.map(c => c.name),
      isSingleGroup: !groupBData
    };

    // Initialize result structure
    timePoints.forEach(timePoint => {
      result.groupA[timePoint] = {};
      if (result.groupB) {
        result.groupB[timePoint] = {};
      }
      
      categories.forEach(category => {
        result.groupA[timePoint][category.name] = { count: 0, percentage: 0 };
        if (result.groupB) {
          result.groupB[timePoint][category.name] = { count: 0, percentage: 0 };
        }
      });
    });

    // Calculate patient-level averages for Group A
    const groupAPatientAverages = this._calculatePatientAverageImprovements(
      commonParams, 
      groupAData.rows,
      'A'
    );
    
    // Calculate patient-level averages for Group B (if available)
    let groupBPatientAverages = null;
    if (groupBData) {
      groupBPatientAverages = this._calculatePatientAverageImprovements(
        commonParams, 
        groupBData.rows,
        'B'
      );
    }

    // Categorize patient averages for Group A
    this._categorizePatientAverages(groupAPatientAverages, result.groupA, categories, 'A');
    
    // Categorize patient averages for Group B (if available)
    if (groupBPatientAverages && result.groupB) {
      this._categorizePatientAverages(groupBPatientAverages, result.groupB, categories, 'B');
    }

    // Calculate percentages
    this._calculateFinalPercentages(result.groupA, timePoints, categories);
    if (result.groupB) {
      this._calculateFinalPercentages(result.groupB, timePoints, categories);
    }

    return result;
  }

  // UPDATED: Calculate patient-level average improvements with null handling for single group
  _calculatePatientAverageImprovements(commonParams, rows, groupName) {
    const timePoints = ['7th', '14th', '21st', '28th'];
    const patientAverages = {
      '7th': [],
      '14th': [], 
      '21st': [],
      '28th': []
    };

    console.log(`Processing ${rows.length} patients for Group ${groupName}`);

    rows.forEach((row, rowIndex) => {
      
      // For each time point, calculate average improvement across all parameters for this patient
      timePoints.forEach((timePoint, timeIndex) => {
        const patientImprovements = []; // Store improvements for all parameters for this patient at this time point
        
        commonParams.forEach(commonParam => {
          // Use the appropriate parameter group based on availability
          const parameterGroup = (groupName === 'A' || !commonParam.groupB) ? 
            commonParam.groupA : commonParam.groupB;
            
          if (!parameterGroup) return; // Skip if parameter group not available
          
          const btValue = row[parameterGroup.btIndex];
          
          if (this._isNumeric(btValue) && parseFloat(btValue) !== 0) {
            const bt = parseFloat(btValue);
            
            // Get the AT value for this time point
            if (timeIndex < parameterGroup.atIndices.length) {
              const atIndex = parameterGroup.atIndices[timeIndex];
              const atValue = row[atIndex];
              
              if (this._isNumeric(atValue)) {
                const at = parseFloat(atValue);
                const improvementPercent = ((bt - at) / bt) * 100;
                patientImprovements.push(improvementPercent);
              }
            }
          }
        });
        
        // Calculate average improvement for this patient at this time point
        if (patientImprovements.length > 0) {
          const averageImprovement = patientImprovements.reduce((sum, imp) => sum + imp, 0) / patientImprovements.length;
          patientAverages[timePoint].push(averageImprovement);
          
          if (rowIndex < 3) { // Debug first few patients
            console.log(`Group ${groupName} Patient ${rowIndex + 1} ${timePoint}: Individual improvements [${patientImprovements.map(p => p.toFixed(1)).join(', ')}] → Average: ${averageImprovement.toFixed(2)}%`);
          }
        }
      });
    });

    // Log summary
    timePoints.forEach(timePoint => {
      console.log(`Group ${groupName} ${timePoint}: ${patientAverages[timePoint].length} patients processed`);
    });

    return patientAverages;
  }

  // Rest of the methods remain the same...
  // (copying the remaining methods as they were)

  // Generate unpaired t-test table if we have exactly 2 groups (Trial vs Control)
  async _generateUnpairedTTestTable(groupData, existingTables) {
    try {
      console.log('Generating unpaired t-test table...');
      
      // Get Trial Group (Sheet1) and Control Group (Sheet2)
      const trialGroup = groupData.groupA;
      const controlGroup = groupData.groupB;
      
      if (!trialGroup || !controlGroup) {
        console.warn('Need at least 2 sheets for unpaired t-test analysis');
        return null;
      }

      // Find the corresponding paired t-test tables to extract the correct means and SDs
      const trialTable = existingTables.find(t => t.sheetNumber === 1 && t.type === 'statistical_analysis');
      const controlTable = existingTables.find(t => t.sheetNumber === 2 && t.type === 'statistical_analysis');

      if (!trialTable || !controlTable) {
        console.warn('Could not find corresponding paired t-test tables for unpaired analysis');
        return null;
      }

      console.log(`Found paired t-test tables for unpaired analysis`);

      // Calculate unpaired t-test statistics using data from paired t-test tables
      const unpairedStats = this._calculateUnpairedTTestFromPairedTables(trialTable, controlTable);

      // Create the unpaired t-test table
      const unpairedTable = {
        title: 'Trial vs Control Group Comparison (Unpaired T-Test)',
        type: 'unpaired_ttest',
        sheetNumber: 998, // Special number to distinguish from regular tables
        sheetName: 'Unpaired T-Test Analysis',
        statistics: unpairedStats,
        tableNumber: 998
      };

      return unpairedTable;

    } catch (error) {
      console.error('Error generating unpaired t-test table:', error);
      return null;
    }
  }

  // Calculate unpaired t-test statistics using data from paired t-test tables
  _calculateUnpairedTTestFromPairedTables(trialTable, controlTable) {
    const allStats = [];

    // Get statistics arrays from both tables
    const trialStats = Array.isArray(trialTable.statistics) ? trialTable.statistics : [trialTable.statistics];
    const controlStats = Array.isArray(controlTable.statistics) ? controlTable.statistics : [controlTable.statistics];

    // Match parameters and assessments between trial and control
    trialStats.forEach(trialStat => {
      const matchingControlStat = controlStats.find(controlStat => 
        this._normalizeParameterName(trialStat.parameter) === this._normalizeParameterName(controlStat.parameter) &&
        trialStat.assessment === controlStat.assessment
      );

      if (matchingControlStat) {
        // Calculate unpaired t-test using the AT means and SDs from paired t-test tables
        const stats = this._calculateUnpairedTTestFromPairedStats(trialStat, matchingControlStat);
        
        allStats.push({
          parameter: trialStat.parameter,
          assessment: trialStat.assessment,
          trialMean: stats.trialMean,
          trialSD: stats.trialSD,
          trialN: stats.trialN,
          controlMean: stats.controlMean,
          controlSD: stats.controlSD,
          controlN: stats.controlN,
          tValue: stats.tValue,
          degreesOfFreedom: stats.df,
          pValue: stats.pValue,
          meanDifference: stats.meanDifference
        });
      }
    });

    return allStats;
  }

  // Calculate unpaired t-test statistics using paired t-test results
  _calculateUnpairedTTestFromPairedStats(trialStat, controlStat) {
    // Use AT means and SDs from the paired t-test tables
    const trialMean = trialStat.meanAT;  // This matches the A.T Mean from Table 27
    const trialSD = trialStat.sdAT;      // This matches the A.T S.D from Table 27
    const trialN = trialStat.n;          // Sample size from trial group

    const controlMean = controlStat.meanAT;  // This matches the A.T Mean from Table 28
    const controlSD = controlStat.sdAT;      // This matches the A.T S.D from Table 28
    const controlN = controlStat.n;          // Sample size from control group

    const meanDiff = trialMean - controlMean;

    // Calculate pooled standard error for unpaired t-test
    const variance1 = Math.pow(trialSD, 2);
    const variance2 = Math.pow(controlSD, 2);
    
    // Pooled standard error for unpaired t-test
    const pooledSE = Math.sqrt((variance1 / trialN) + (variance2 / controlN));

    // Calculate t-value
    const t = pooledSE === 0 ? (meanDiff === 0 ? 0 : (meanDiff > 0 ? Infinity : -Infinity)) : meanDiff / pooledSE;

    // Degrees of freedom for unpaired t-test: n1 + n2 - 2
    const df = trialN + controlN - 2;

    // Calculate p-value
    const p = this._twoTailedPValue(t, df);

    return {
      trialMean: trialMean,
      trialSD: trialSD,
      trialN: trialN,
      controlMean: controlMean,
      controlSD: controlSD,
      controlN: controlN,
      tValue: t,
      df: df,
      pValue: p,
      meanDifference: meanDiff
    };
  }

  // Find common parameters between two groups
  _findCommonParameters(groupAParams, groupBParams) {
    const commonParams = [];
    
    groupAParams.forEach(paramA => {
      const matchingParamB = groupBParams.find(paramB => 
        this._normalizeParameterName(paramA.name) === this._normalizeParameterName(paramB.name)
      );
      
      if (matchingParamB) {
        commonParams.push({
          name: paramA.name,
          groupA: paramA,
          groupB: matchingParamB
        });
      }
    });

    return commonParams;
  }

  // Normalize parameter names for comparison
  _normalizeParameterName(name) {
    return name.toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  // Categorize patient average improvements
  _categorizePatientAverages(patientAverages, groupResult, categories, groupName) {
    Object.keys(patientAverages).forEach(timePoint => {
      const averageValues = patientAverages[timePoint];
      
      console.log(`Categorizing Group ${groupName} ${timePoint}: ${averageValues.length} patient averages`);
      
      averageValues.forEach((averagePercent, patientIndex) => {
        let categorized = false;
        
        categories.forEach(category => {
          if (categorized) return; // Skip if already categorized
          
          const min = category.min === -Infinity ? -Infinity : category.min;
          const max = category.max === Infinity ? Infinity : category.max;
          
          // Inclusive of min, exclusive of max (except for the top category)
          const inRange = (max === Infinity) ? 
            (averagePercent >= min) : 
            (averagePercent >= min && averagePercent < max);
          
          if (inRange) {
            groupResult[timePoint][category.name].count++;
            categorized = true;
            
            if (patientIndex < 3) { // Debug first few patients
              console.log(`  Patient ${patientIndex + 1}: ${averagePercent.toFixed(2)}% → ${category.name}`);
            }
          }
        });
        
        if (!categorized) {
          console.warn(`Patient average ${averagePercent.toFixed(2)}% could not be categorized for ${timePoint}`);
        }
      });
      
      // Log category totals
      const totalCount = Object.values(groupResult[timePoint]).reduce((sum, cat) => sum + cat.count, 0);
      console.log(`Group ${groupName} ${timePoint} total categorized: ${totalCount} patients`);
    });
  }

  // Calculate final percentages (totals are already correct)
  _calculateFinalPercentages(groupResult, timePoints, categories) {
    timePoints.forEach(timePoint => {
      let total = 0;
      
      // Calculate total count for this time point
      categories.forEach(category => {
        total += groupResult[timePoint][category.name].count;
      });
      
      // Calculate percentages
      categories.forEach(category => {
        const count = groupResult[timePoint][category.name].count;
        const percentage = total > 0 ? Math.round((count / total) * 100) : 0;
        groupResult[timePoint][category.name].percentage = percentage;
      });
    });
  }

  // Rest of existing methods remain unchanged...
  async _extractSheetDataWithMultiRowHeaders(worksheet) {
    const maxHeaderRows = 5;
    const headerRows = [];
    const data = { headers: [], rows: [] };

    for (let r = 1; r <= Math.min(maxHeaderRows, worksheet.actualRowCount); r++) {
      const row = worksheet.getRow(r);
      const vals = [];
      for (let c = 1; c <= worksheet.columnCount; c++) {
        const v = row.getCell(c).value;
        vals.push((v === null || v === undefined) ? '' : String(v).trim());
      }
      headerRows.push(vals);
      const numericCount = vals.filter(x => x !== '' && !isNaN(parseFloat(x))).length;
      if (numericCount > Math.floor(worksheet.columnCount / 2)) break;
    }

    const row2 = headerRows[1] || [];
    const row3 = headerRows[2] || [];
    const row2Text = row2.join(' ').toLowerCase();
    const row3Text = row3.join(' ').toLowerCase();

    const row2HasBTAT = /(^|\W)(bt|at|before|after)(\W|$)/i.test(row2Text);
    const row3HasDay = /\b(7|14|21|28|7th|14th|21st|28th|day|d)\b/i.test(row3Text);

    let headerRowCount = 1;
    if (row2HasBTAT && row3HasDay) headerRowCount = 3;
    else if (row2HasBTAT) headerRowCount = 2;

    const composed = [];
    for (let c = 1; c <= worksheet.columnCount; c++) {
      const parts = [];
      for (let r = 1; r <= headerRowCount; r++) {
        const cell = worksheet.getRow(r).getCell(c).value;
        if (cell !== null && cell !== undefined && String(cell).trim() !== '') parts.push(String(cell).trim());
      }
      composed.push(parts.join(' | ').trim());
    }

    data.headers = composed;

    for (let r = headerRowCount + 1; r <= worksheet.actualRowCount; r++) {
      const row = worksheet.getRow(r);
      const rowData = [];
      let hasData = false;
      for (let c = 1; c <= worksheet.columnCount; c++) {
        const v = row.getCell(c).value;
        if (v !== null && v !== undefined && String(v).trim() !== '') hasData = true;
        rowData[c - 1] = v;
      }
      if (hasData) data.rows.push(rowData);
    }

    return data;
  }

  _hasValidData(sheetData) {
    if (!sheetData || !sheetData.headers || !sheetData.rows) return false;
    if (sheetData.headers.length < 2) return false;
    if (sheetData.rows.length === 0) return false;
    return true;
  }

  _identifyParameterGroups(headers) {
    const groups = [];
    const skipTokens = ['sl','sl.','sl.no','slno','no','patient','id'];

    const parsed = headers.map((h,idx) => {
      const parts = (h || '').split('|').map(p => p.trim()).filter(Boolean);
      const top = parts.length ? parts[0] : '';
      const lowerParts = parts.map(p => p.toLowerCase());
      return { index: idx, raw: h || '', parts, top: top.trim(), lowerParts };
    });

    const map = new Map();
    parsed.forEach(p => {
      const topClean = (p.top || '').toString().toLowerCase().replace(/[^a-z0-9]/g,'');
      if (!topClean) return;
      if (skipTokens.some(t => topClean.includes(t))) return;
      if (!map.has(topClean)) map.set(topClean, []);
      map.get(topClean).push(p);
    });

    for (const [k, cols] of map.entries()) {
      let btCol = null;
      const atCols = [];

      for (const p of cols) {
        if (p.lowerParts.some(lp => /(^|\W)(bt|before|b\.t)(\W|$)/.test(lp))) {
          btCol = p; break;
        }
      }
      if (!btCol) {
        btCol = cols.find(p => !p.lowerParts.some(lp => lp.includes('at') || lp.includes('after') || /\b(7|14|21|28|day|d)\b/.test(lp))) || cols[0];
      }

      for (const p of cols) {
        if (p.index === btCol.index) continue;
        atCols.push(p);
      }
      atCols.sort((a,b)=>a.index-b.index);
      if (!btCol || atCols.length === 0) continue;

      groups.push({
        name: (btCol.top || k).toString().trim(),
        btIndex: btCol.index,
        btHeader: btCol.raw,
        atIndices: atCols.map(c => c.index),
        atHeaders: atCols.map(c => c.raw)
      });
    }

    return groups;
  }

  async _processSheetIntoSingleTable(sheetData, sheetNumber, sheetName) {
    const headers = sheetData.headers;
    const parameterGroups = this._identifyParameterGroups(headers);
    const allStats = [];

    for (const group of parameterGroups) {
      for (let k = 0; k < group.atIndices.length; k++) {
        const atIdx = group.atIndices[k];
        const atHeader = group.atHeaders ? group.atHeaders[k] : '';
        const btValues = [];
        const atValues = [];

        for (const row of sheetData.rows) {
          const btVal = row[group.btIndex];
          const atVal = row[atIdx];
          if (this._isNumeric(btVal) && this._isNumeric(atVal)) {
            btValues.push(parseFloat(btVal));
            atValues.push(parseFloat(atVal));
          }
        }

        if (btValues.length === 0 || atValues.length === 0) continue;

        const stats = this._calculateStatistics(btValues, atValues);
        allStats.push({
          parameter: group.name,
          assessment: this._deriveAssessmentLabel(atHeader, k),
          ...stats
        });
      }
    }

    if (allStats.length === 0) return null;

    return {
      title: `Master Chart - ${sheetName}`,
      type: 'statistical_analysis',
      sheetNumber,
      sheetName,
      statistics: allStats,
      tableNumber: sheetNumber
    };
  }

  _deriveAssessmentLabel(header, orderIndex) {
    if (header) {
      const h = header.toString().toLowerCase();
      const dayMatch = h.match(/\b(7|14|21|28)\b/);
      if (dayMatch) return `${dayMatch[1]}th day`;
      const ord = h.match(/\b(\d+)(st|nd|rd|th)\b/);
      if (ord) return `${ord[1]}th day`;
      if (h.includes('day')) return header.toString().trim();
    }
    const map = ['7th day','14th day','21st day','28th day'];
    return map[orderIndex] || `Assessment ${orderIndex+1}`;
  }

  _isNumeric(v) {
    if (v === null || v === undefined || v === '') return false;
    const n = parseFloat(v);
    return !isNaN(n) && isFinite(n);
  }

  _calculateStatistics(btValues, atValues) {
    const n = btValues.length;
    if (n !== atValues.length) throw new Error('paired arrays length mismatch');

    const diffs = btValues.map((b,i) => b - atValues[i]);
    const meanBT = this._mean(btValues);
    const meanAT = this._mean(atValues);
    const meanDiff = this._mean(diffs);
    const sdBT = this._stdDev(btValues);
    const sdAT = this._stdDev(atValues);
    const sdDiff = this._stdDev(diffs);
    const se = sdDiff / Math.sqrt(n);
    const t = se === 0 ? (meanDiff === 0 ? 0 : (meanDiff > 0 ? Infinity : -Infinity)) : meanDiff / se;
    const df = Math.max(0, n - 1);
    const p = this._twoTailedPValue(t, df);
    const eff = (meanBT === 0) ? 0 : (meanDiff / meanBT) * 100;

    return { 
      n, 
      meanBT, 
      meanAT, 
      meanDiff, 
      sdBT, 
      sdAT, 
      sdDiff, 
      standardError: se, 
      tValue: t, 
      degreesOfFreedom: df, 
      pValue: p,
      effectiveness: eff, 
      rawDifferences: diffs 
    };
  }

  _mean(arr) { if (!arr || arr.length===0) return 0; return arr.reduce((s,v)=>s+v,0)/arr.length; }
  _stdDev(arr) { if (!arr || arr.length<2) return 0; const m = this._mean(arr); const sq = arr.map(v => Math.pow(v-m,2)); const varr = sq.reduce((s,v)=>s+v,0)/(arr.length-1); return Math.sqrt(varr); }

  _twoTailedPValue(tValue, df) {
    if (!isFinite(tValue)) return tValue === 0 ? 1 : 0;
    const absT = Math.abs(tValue);
    if (df > 30) {
      const z = absT;
      return Math.max(0, Math.min(1, 2 * 0.5 * (1 - this._erf(z / Math.sqrt(2)))));
    }
    const p = 2 * (1 - this._normalCDF(absT));
    return Math.max(0, Math.min(1, p));
  }

  _normalCDF(x) { return 0.5 * (1 + this._erf(x / Math.sqrt(2))); }
  _erf(x) {
    const a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
    const sign = x < 0 ? -1 : 1;
    x = Math.abs(x);
    const t = 1 / (1 + p * x);
    const y = 1 - ((((a5*t + a4)*t + a3)*t + a2)*t + a1) * t * Math.exp(-x*x);
    return sign * y;
  }
}

module.exports = new MasterChartService();