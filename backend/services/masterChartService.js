// backend/services/masterChartService.js
const ExcelJS = require('exceljs');

class MasterChartService {
  async processMasterChart(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const processedTables = [];

    for (let sheetIndex = 0; sheetIndex < workbook.worksheets.length; sheetIndex++) {
      const worksheet = workbook.worksheets[sheetIndex];
      if (!worksheet || worksheet.actualRowCount === 0) continue;

      const sheetData = await this._extractSheetDataWithMultiRowHeaders(worksheet);
      if (!this._hasValidData(sheetData)) continue;

      const table = await this._processSheetIntoSingleTable(sheetData, sheetIndex + 1, worksheet.name);
      if (table) processedTables.push(table);
    }

    if (processedTables.length === 0) throw new Error('No valid data found in master chart.');

    return {
      tables: processedTables,
      sourceFile: filePath,
      totalSheets: workbook.worksheets.length,
      processedSheets: processedTables.length
    };
  }

  // --- header extraction (1-3 rows) ---
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

  // --- group identification ---
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

  // --- SINGLE TABLE CREATION ---
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
      pValue: p,   // return raw p-value directly
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
