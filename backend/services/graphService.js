// /* graphService.updated.js
//    - platform width (depth) increased while keeping X length same
//    - zero gap between group bars (trial/control/total), group spacing between categories unchanged
//    (other behavior unchanged)
// */

// const { createCanvas, registerFont } = require('canvas');
// const Chart = require('chart.js/auto');
// const fs = require('fs').promises;
// const fsSync = require('fs');
// const path = require('path');

// const UPLOAD_DIR = path.join(process.cwd(), 'uploads');

// class GraphService {
//   constructor() { this.ensureUploadDir(); }
//   async ensureUploadDir() {
//     try {
//       if (!fsSync.existsSync(UPLOAD_DIR)) {
//         await fs.mkdir(UPLOAD_DIR, { recursive: true });
//       }
//     } catch (err) {
//       console.warn('Could not create uploads directory:', err.message);
//     }
//   }

//   shade(hex, percent) {
//     const num = parseInt(hex.slice(1), 16);
//     let r = (num >> 16) + Math.round(255 * percent);
//     let g = ((num >> 8) & 0x00FF) + Math.round(255 * percent);
//     let b = (num & 0x0000FF) + Math.round(255 * percent);
//     r = Math.max(0, Math.min(255, r));
//     g = Math.max(0, Math.min(255, g));
//     b = Math.max(0, Math.min(255, b));
//     return `rgb(${r}, ${g}, ${b})`;
//   }

//   async generateGraphs(tables) {
//     if (!Array.isArray(tables)) throw new Error('tables must be an array');
//     const graphs = [];
//     for (let i = 0; i < tables.length; i++) {
//       const table = tables[i];
//       const imagePath = await this.createSPSS3DBarChart(table, i + 1).catch(async err => {
//         console.error('3D rect chart failed:', err);
//         return await this.createHighQualityBarChart(table, i + 1);
//       });
//       graphs.push({
//         title: table.title || `Table ${i + 1}`,
//         tableNumber: i + 1,
//         imagePath
//       });
//     }
//     return graphs;
//   }

//   async createSPSS3DBarChart(tableData, tableNumber) {
//     const width = 2400;
//     const height = 1500;
//     const canvas = createCanvas(width, height);
//     const ctx = canvas.getContext('2d');
//     ctx.imageSmoothingEnabled = true;
//     ctx.imageSmoothingQuality = 'high';

//     if (!tableData || !Array.isArray(tableData.rows)) {
//       throw new Error('Invalid tableData.rows');
//     }

//     const dataRows = tableData.rows.filter(r => String(r.category || '').toLowerCase() !== 'total');
//     const labels = dataRows.map(r => String(r.category || ''));
//     const trialData = dataRows.map(r => Number(r.trialGroup) || 0);
//     const controlData = dataRows.map(r => Number(r.controlGroup) || 0);
//     const totalData = dataRows.map(r => Number(r.total) || 0);

//     const palette = Object.assign({
//       trial: '#9fb4ffff',
//       control: '#8e4b6eff',
//       total: '#efe6b2ff'
//     }, tableData.colors || {});

//     const xAxisTitle = tableData.xLabel || tableData.xTitle || (tableData.title && tableData.title.split('-').pop().trim()) || 'Category';
//     const yAxisTitle = tableData.yLabel || tableData.yTitle || 'Grades';

//     const chart = new Chart(ctx, {
//       type: 'bar',
//       data: {
//         labels,
//         datasets: [
//           { label: 'Trial Group', data: trialData, backgroundColor: palette.trial, borderColor: '#000000ff', borderWidth: 2, barPercentage: 0.5, categoryPercentage: 0.6 },
//           { label: 'Control Group', data: controlData, backgroundColor: palette.control, borderColor: '#000000ff', borderWidth: 2, barPercentage: 0.5, categoryPercentage: 0.6 },
//           { label: 'Total', data: totalData, backgroundColor: palette.total, borderColor: '#000000ff', borderWidth: 2, barPercentage: 0.5, categoryPercentage: 0.6 }
//         ]
//       },
//       options: {
//         responsive: false,
//         animation: false,
//         maintainAspectRatio: false,
//         layout: { padding: { left: 120, right: 260, top: 90, bottom: 140 } },
//         plugins: {
//           title: {
//             display: true,
//             text: `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''}`,
//             color: '#111827',
//             font: { size: 36, weight: '700', family: "'Helvetica Neue', Arial" },
//             padding: 70
//           },
//           legend: {
//             display: true,
//             position: 'right',
//             align: 'center',
//             labels: { font: { size: 20, weight: '600' } }
//           },
//           tooltip: { enabled: true }
//         },
//         scales: {
//           x: {
//             ticks: { font: { size: 24, weight: '600' }, color: '#000000ff' },
//             grid: { display: false, drawBorder: true, borderWidth: 4, borderColor: '#111827' },
//             title: { display: true, text: xAxisTitle, font: { size: 30, weight: '800' }, color: '#000000ff' }
//           },
//           y: {
//             beginAtZero: true,
//             ticks: { font: { size: 24, weight: '600' }, stepSize: 2, color: '#000000ff' },
//             grid: { display: false, color: 'rgba(0,0,0,0.06)', lineWidth: 1 },
//             title: { display: true, text: yAxisTitle, font: { size: 30, weight: '800' }, color: '#000000ff' }
//           }
//         },
//         elements: { bar: { borderWidth: 5 } }
//       },
//       plugins: [
//         {
//           id: 'spss-3d-rect-platform',
//           beforeDatasetsDraw: chartInstance => {
//             const ctx = chartInstance.ctx;
//             const chartArea = chartInstance.chartArea;
//             const xScale = chartInstance.scales.x;

//             if (!xScale || !chartInstance.data.labels.length) return;

//             // base dx/dy used to draw bars (keep these for bars)
//             const dx = Math.max(8, Math.round((chartArea.right - chartArea.left) * 0.01));
//             const dy = Math.max(6, Math.round(dx * 0.45));

//             // PLATFORM: increase depth (width from front to back) while keeping X length same
//             // Use larger offsets for the platform only so it appears wider in perspective
//             const platformDx = Math.round(dx * 7);  // increased depth multiplier
//             const platformDy = Math.round(dy * 6);

//             const yScale = chartInstance.scales.y;
//             const baseY = (yScale && typeof yScale.getPixelForValue === 'function') ? yScale.getPixelForValue(0) : chartArea.bottom;

//             ctx.save();
//             ctx.beginPath();
//             ctx.fillStyle = '#9aa0a6'; // top surface grey
//             ctx.strokeStyle = '#4b4f52';
//             ctx.lineWidth = 2.2;

//             const leftX = chartArea.left - dx - 6;
//             const rightX = chartArea.right + 6;
//             // Note: keep leftX/rightX as before (same X length), but use platformDx/platformDy for perspective
//             ctx.moveTo(leftX, baseY);
//             ctx.lineTo(rightX, baseY);
//             ctx.lineTo(rightX + platformDx, baseY - platformDy);
//             ctx.lineTo(leftX + platformDx, baseY - platformDy);
//             ctx.closePath();
//             ctx.fill();
//             ctx.stroke();

//             // subtle top highlight
//             ctx.beginPath();
//             ctx.fillStyle = 'rgba(255,255,255,0.06)';
//             ctx.moveTo(leftX + 6, baseY - 2);
//             ctx.lineTo(rightX - 6, baseY - 2);
//             ctx.lineTo(rightX - 6 + platformDx, baseY - platformDy - 2);
//             ctx.lineTo(leftX + 6 + platformDx, baseY - platformDy - 2);
//             ctx.closePath();
//             ctx.fill();

//             ctx.restore();
//           },

//           afterDatasetsDraw: chartInstance => {
//             const ctx = chartInstance.ctx;
//             const chartArea = chartInstance.chartArea;
//             const datasets = chartInstance.data.datasets;
//             const xScale = chartInstance.scales.x;

//             // compute dx/dy for bars (unchanged -- bars keep original offsets)
//             const dx = Math.max(8, Math.round((chartArea.right - chartArea.left) * 0.01));
//             const dy = Math.max(6, Math.round(dx * 0.45));

//             // We'll draw grouped bars ourselves but preserve bar centers from Chart.js.
//             // To make the three bars in each category touch (gap=0) we'll:
//             //  - compute category span (minLeft, maxRight) across datasets for this category index
//             //  - split that span equally among the number of datasets (so bars abut)
//             const datasetCount = datasets.length;
//             const metas = datasets.map((_, di) => chartInstance.getDatasetMeta(di));

//             // number of categories:
//             const categoryCount = (metas[0] && metas[0].data) ? metas[0].data.length : 0;

//             for (let idx = 0; idx < categoryCount; idx++) {
//               // collect left/right from each dataset's bar for this category idx
//               const lefts = [];
//               const rights = [];
//               const barsForCategory = [];

//               for (let di = 0; di < datasetCount; di++) {
//                 const meta = metas[di];
//                 if (!meta || !meta.data || !meta.data[idx]) continue;
//                 const bar = meta.data[idx];
//                 barsForCategory.push({ bar, dsIndex: di });
//                 const barLeft = bar.x - (bar.width / 2);
//                 const barRight = bar.x + (bar.width / 2);
//                 lefts.push(barLeft);
//                 rights.push(barRight);
//               }

//               if (barsForCategory.length === 0) continue;

//               const minLeft = Math.min(...lefts);
//               const maxRight = Math.max(...rights);
//               const totalSpan = Math.max(1, maxRight - minLeft);

//               const newWidth = totalSpan / barsForCategory.length;

//               // draw each dataset's bar using computed equal subdivisions so gaps = 0
//               barsForCategory.forEach((entry, placeIndex) => {
//                 const bar = entry.bar;
//                 const dsIndex = entry.dsIndex;

//                 ctx.save();

//                 const leftX = minLeft + placeIndex * newWidth;
//                 const rightX = leftX + newWidth;
//                 const w = Math.max(1, rightX - leftX);
//                 const xCenter = leftX + w / 2;

//                 const y = bar.y;
//                 const bottomY = bar.base;

//                 const ds = datasets[dsIndex];
//                 const baseColor = ds.backgroundColor || '#777';

//                 // FRONT FACE
//                 ctx.fillStyle = baseColor;
//                 ctx.beginPath();
//                 ctx.rect(leftX, y, w, bottomY - y);
//                 ctx.fill();

//                 // TOP FACE (use original dx/dy for bars to preserve perspective)
//                 ctx.fillStyle = baseColor;
//                 ctx.beginPath();
//                 ctx.moveTo(leftX, y);
//                 ctx.lineTo(rightX, y);
//                 ctx.lineTo(rightX + dx, y - dy);
//                 ctx.lineTo(leftX + dx, y - dy);
//                 ctx.closePath();
//                 ctx.fill();

//                 // RIGHT SIDE FACE
//                 ctx.fillStyle = baseColor;
//                 ctx.beginPath();
//                 ctx.moveTo(rightX, y);
//                 ctx.lineTo(rightX + dx, y - dy);
//                 ctx.lineTo(rightX + dx, bottomY - dy);
//                 ctx.lineTo(rightX, bottomY);
//                 ctx.closePath();
//                 ctx.fill();

//                 // outlines
//                 ctx.lineWidth = 2.4;
//                 ctx.strokeStyle = '#111827';

//                 ctx.beginPath();
//                 ctx.rect(leftX, y, w, bottomY - y);
//                 ctx.stroke();

//                 ctx.beginPath();
//                 ctx.moveTo(leftX, y);
//                 ctx.lineTo(rightX, y);
//                 ctx.lineTo(rightX + dx, y - dy);
//                 ctx.lineTo(leftX + dx, y - dy);
//                 ctx.closePath();
//                 ctx.stroke();

//                 ctx.beginPath();
//                 ctx.moveTo(rightX, y);
//                 ctx.lineTo(rightX + dx, y - dy);
//                 ctx.lineTo(rightX + dx, bottomY - dy);
//                 ctx.lineTo(rightX, bottomY);
//                 ctx.closePath();
//                 ctx.stroke();

//                 // numeric label above top (perspective shifted)
//                 const value = ds.data[idx];
//                 if (typeof value !== 'undefined' && value !== null) {
//                   const labelX = xCenter + dx / 2;
//                   const labelY = y - dy - 2;
//                   ctx.font = '700 28px "Helvetica Neue", Arial';
//                   ctx.fillStyle = '#111827';
//                   ctx.textAlign = 'center';
//                   ctx.textBaseline = 'bottom';
//                   ctx.fillText(String(value), labelX, labelY);
//                 }

//                 ctx.restore();
//               });
//             } // end category loop
//           } // end afterDatasetsDraw
//         } // end plugin
//       ] // end plugins
//     });

//     chart.update();

//     await this.ensureUploadDir();
//     const filename = `graph-3drect-${tableNumber}-${Date.now()}.png`;
//     const filepath = path.join(UPLOAD_DIR, filename);
//     const buffer = canvas.toBuffer('image/png');
//     await fs.writeFile(filepath, buffer);

//     try { chart.destroy(); } catch (e) { /* ignore */ }

//     return filepath;
//   }

//   // fallback unchanged
//   async createHighQualityBarChart(tableData, tableNumber) {
//     const width = 1600;
//     const height = 1000;
//     const canvas = createCanvas(width, height);
//     const ctx = canvas.getContext('2d');
//     ctx.imageSmoothingEnabled = true;
//     ctx.imageSmoothingQuality = 'high';

//     if (!tableData || !Array.isArray(tableData.rows)) {
//       throw new Error('Invalid tableData.rows');
//     }

//     const dataRows = tableData.rows.filter(r => String(r.category || '').toLowerCase() !== 'total');
//     const labels = dataRows.map(r => String(r.category || ''));
//     const trialData = dataRows.map(r => Number(r.trialGroup) || 0);
//     const controlData = dataRows.map(r => Number(r.controlGroup) || 0);
//     const totalData = dataRows.map(r => Number(r.total) || 0);

//     const colors = {
//       trial: { background: 'rgba(123,94,126,0.95)', border: 'rgba(90,60,90,1)' },
//       control: { background: 'rgba(143,160,255,0.95)', border: 'rgba(100,120,210,1)' },
//       total: { background: 'rgba(245,240,184,0.98)', border: 'rgba(220,200,100,1)' }
//     };

//     const chart = new Chart(ctx, {
//       type: 'bar',
//       data: {
//         labels,
//         datasets: [
//           { label: 'Trial Group', data: trialData, backgroundColor: colors.trial.background, borderColor: colors.trial.border, borderWidth: 2 },
//           { label: 'Control Group', data: controlData, backgroundColor: colors.control.background, borderColor: colors.control.border, borderWidth: 2 },
//           { label: 'Total', data: totalData, backgroundColor: colors.total.background, borderColor: colors.total.border, borderWidth: 2 }
//         ]
//       },
//       options: {
//         responsive: false,
//         animation: false,
//         maintainAspectRatio: false,
//         plugins: {
//           title: {
//             display: true,
//             text: `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''}`,
//             font: { size: 28, weight: '700' },
//             padding: 20
//           },
//           legend: { display: true, position: 'top' }
//         },
//         scales: {
//           x: { ticks: { font: { size: 16, weight: '700' } }, grid: { display: false } },
//           y: { beginAtZero: true, ticks: { font: { size: 16, weight: '700' }, stepSize: 1 } }
//         }
//       }
//     });

//     chart.update();

//     await this.ensureUploadDir();
//     const filename = `graph-hq-${tableNumber}-${Date.now()}.png`;
//     const filepath = path.join(UPLOAD_DIR, filename);
//     const buffer = canvas.toBuffer('image/png');
//     await fs.writeFile(filepath, buffer);

//     try { chart.destroy(); } catch (_) { }
//     return filepath;
//   }
// }

// module.exports = new GraphService();



/* graphService.js - Updated for PNG export and Y-axis label fix
   - Changed Y-axis label from "Grades" to "Number of patients (No. patients)"
   - Maintained all existing 3D bar chart functionality
   - Ready for ZIP file generation workflow
*/

const { createCanvas, registerFont } = require('canvas');
const Chart = require('chart.js/auto');
const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');

const UPLOAD_DIR = path.join(process.cwd(), 'uploads');

class GraphService {
  constructor() { this.ensureUploadDir(); }
  async ensureUploadDir() {
    try {
      if (!fsSync.existsSync(UPLOAD_DIR)) {
        await fs.mkdir(UPLOAD_DIR, { recursive: true });
      }
    } catch (err) {
      console.warn('Could not create uploads directory:', err.message);
    }
  }

  shade(hex, percent) {
    const num = parseInt(hex.slice(1), 16);
    let r = (num >> 16) + Math.round(255 * percent);
    let g = ((num >> 8) & 0x00FF) + Math.round(255 * percent);
    let b = (num & 0x0000FF) + Math.round(255 * percent);
    r = Math.max(0, Math.min(255, r));
    g = Math.max(0, Math.min(255, g));
    b = Math.max(0, Math.min(255, b));
    return `rgb(${r}, ${g}, ${b})`;
  }

  async generateGraphs(tables) {
    if (!Array.isArray(tables)) throw new Error('tables must be an array');
    const graphs = [];
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      const imagePath = await this.createSPSS3DBarChart(table, i + 1).catch(async err => {
        console.error('3D rect chart failed:', err);
        return await this.createHighQualityBarChart(table, i + 1);
      });
      graphs.push({
        title: table.title || `Table ${i + 1}`,
        tableNumber: i + 1,
        imagePath,
        filename: `graph-${i + 1}-${this.sanitizeFilename(table.title || 'table')}.png`
      });
    }
    return graphs;
  }

  // Helper method to sanitize filename
  sanitizeFilename(text) {
    return text.replace(/[^a-zA-Z0-9-_]/g, '-').toLowerCase();
  }

  async createSPSS3DBarChart(tableData, tableNumber) {
    const width = 2400;
    const height = 1500;
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';

    if (!tableData || !Array.isArray(tableData.rows)) {
      throw new Error('Invalid tableData.rows');
    }

    const dataRows = tableData.rows.filter(r => String(r.category || '').toLowerCase() !== 'total');
    const labels = dataRows.map(r => String(r.category || ''));
    const trialData = dataRows.map(r => Number(r.trialGroup) || 0);
    const controlData = dataRows.map(r => Number(r.controlGroup) || 0);
    const totalData = dataRows.map(r => Number(r.total) || 0);

    const palette = Object.assign({
      trial: '#9fb4ffff',
      control: '#8e4b6eff',
      total: '#efe6b2ff'
    }, tableData.colors || {});

    const xAxisTitle = tableData.xLabel || tableData.xTitle || (tableData.title && tableData.title.split('-').pop().trim()) || 'Category';
    // FIXED: Changed Y-axis label from "Grades" to "Number of patients (No. patients)"
    const yAxisTitle = tableData.yLabel || tableData.yTitle || 'Number of patients (No. patients)';

    const chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Trial Group', data: trialData, backgroundColor: palette.trial, borderColor: '#000000ff', borderWidth: 2, barPercentage: 0.5, categoryPercentage: 0.6 },
          { label: 'Control Group', data: controlData, backgroundColor: palette.control, borderColor: '#000000ff', borderWidth: 2, barPercentage: 0.5, categoryPercentage: 0.6 },
          { label: 'Total', data: totalData, backgroundColor: palette.total, borderColor: '#000000ff', borderWidth: 2, barPercentage: 0.5, categoryPercentage: 0.6 }
        ]
      },
      options: {
        responsive: false,
        animation: false,
        maintainAspectRatio: false,
        layout: { padding: { left: 120, right: 260, top: 90, bottom: 140 } },
        plugins: {
          title: {
            display: true,
            text: `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''}`,
            color: '#111827',
            font: { size: 36, weight: '700', family: "'Helvetica Neue', Arial" },
            padding: 70
          },
          legend: {
            display: true,
            position: 'right',
            align: 'center',
            labels: { font: { size: 20, weight: '600' } }
          },
          tooltip: { enabled: true }
        },
        scales: {
          x: {
            ticks: { font: { size: 24, weight: '600' }, color: '#000000ff' },
            grid: { display: false, drawBorder: true, borderWidth: 4, borderColor: '#111827' },
            title: { display: true, text: xAxisTitle, font: { size: 30, weight: '800' }, color: '#000000ff' }
          },
          y: {
            beginAtZero: true,
            ticks: { font: { size: 24, weight: '600' }, stepSize: 2, color: '#000000ff' },
            grid: { display: false, color: 'rgba(0,0,0,0.06)', lineWidth: 1 },
            title: { display: true, text: yAxisTitle, font: { size: 30, weight: '800' }, color: '#000000ff' }
          }
        },
        elements: { bar: { borderWidth: 5 } }
      },
      plugins: [
        {
          id: 'spss-3d-rect-platform',
          beforeDatasetsDraw: chartInstance => {
            const ctx = chartInstance.ctx;
            const chartArea = chartInstance.chartArea;
            const xScale = chartInstance.scales.x;

            if (!xScale || !chartInstance.data.labels.length) return;

            const dx = Math.max(8, Math.round((chartArea.right - chartArea.left) * 0.01));
            const dy = Math.max(6, Math.round(dx * 0.45));
            const platformDx = Math.round(dx * 7);
            const platformDy = Math.round(dy * 6);

            const yScale = chartInstance.scales.y;
            const baseY = (yScale && typeof yScale.getPixelForValue === 'function') ? yScale.getPixelForValue(0) : chartArea.bottom;

            ctx.save();
            ctx.beginPath();
            ctx.fillStyle = '#9aa0a6';
            ctx.strokeStyle = '#4b4f52';
            ctx.lineWidth = 2.2;

            const leftX = chartArea.left - dx - 6;
            const rightX = chartArea.right + 6;
            ctx.moveTo(leftX, baseY);
            ctx.lineTo(rightX, baseY);
            ctx.lineTo(rightX + platformDx, baseY - platformDy);
            ctx.lineTo(leftX + platformDx, baseY - platformDy);
            ctx.closePath();
            ctx.fill();
            ctx.stroke();

            ctx.beginPath();
            ctx.fillStyle = 'rgba(255,255,255,0.06)';
            ctx.moveTo(leftX + 6, baseY - 2);
            ctx.lineTo(rightX - 6, baseY - 2);
            ctx.lineTo(rightX - 6 + platformDx, baseY - platformDy - 2);
            ctx.lineTo(leftX + 6 + platformDx, baseY - platformDy - 2);
            ctx.closePath();
            ctx.fill();

            ctx.restore();
          },

          afterDatasetsDraw: chartInstance => {
            const ctx = chartInstance.ctx;
            const chartArea = chartInstance.chartArea;
            const datasets = chartInstance.data.datasets;
            const xScale = chartInstance.scales.x;

            const dx = Math.max(8, Math.round((chartArea.right - chartArea.left) * 0.01));
            const dy = Math.max(6, Math.round(dx * 0.45));

            const datasetCount = datasets.length;
            const metas = datasets.map((_, di) => chartInstance.getDatasetMeta(di));
            const categoryCount = (metas[0] && metas[0].data) ? metas[0].data.length : 0;

            for (let idx = 0; idx < categoryCount; idx++) {
              const lefts = [];
              const rights = [];
              const barsForCategory = [];

              for (let di = 0; di < datasetCount; di++) {
                const meta = metas[di];
                if (!meta || !meta.data || !meta.data[idx]) continue;
                const bar = meta.data[idx];
                barsForCategory.push({ bar, dsIndex: di });
                const barLeft = bar.x - (bar.width / 2);
                const barRight = bar.x + (bar.width / 2);
                lefts.push(barLeft);
                rights.push(barRight);
              }

              if (barsForCategory.length === 0) continue;

              const minLeft = Math.min(...lefts);
              const maxRight = Math.max(...rights);
              const totalSpan = Math.max(1, maxRight - minLeft);
              const newWidth = totalSpan / barsForCategory.length;

              barsForCategory.forEach((entry, placeIndex) => {
                const bar = entry.bar;
                const dsIndex = entry.dsIndex;

                ctx.save();

                const leftX = minLeft + placeIndex * newWidth;
                const rightX = leftX + newWidth;
                const w = Math.max(1, rightX - leftX);
                const xCenter = leftX + w / 2;

                const y = bar.y;
                const bottomY = bar.base;

                const ds = datasets[dsIndex];
                const baseColor = ds.backgroundColor || '#777';

                // FRONT FACE
                ctx.fillStyle = baseColor;
                ctx.beginPath();
                ctx.rect(leftX, y, w, bottomY - y);
                ctx.fill();

                // TOP FACE
                ctx.fillStyle = baseColor;
                ctx.beginPath();
                ctx.moveTo(leftX, y);
                ctx.lineTo(rightX, y);
                ctx.lineTo(rightX + dx, y - dy);
                ctx.lineTo(leftX + dx, y - dy);
                ctx.closePath();
                ctx.fill();

                // RIGHT SIDE FACE
                ctx.fillStyle = baseColor;
                ctx.beginPath();
                ctx.moveTo(rightX, y);
                ctx.lineTo(rightX + dx, y - dy);
                ctx.lineTo(rightX + dx, bottomY - dy);
                ctx.lineTo(rightX, bottomY);
                ctx.closePath();
                ctx.fill();

                // outlines
                ctx.lineWidth = 2.4;
                ctx.strokeStyle = '#111827';

                ctx.beginPath();
                ctx.rect(leftX, y, w, bottomY - y);
                ctx.stroke();

                ctx.beginPath();
                ctx.moveTo(leftX, y);
                ctx.lineTo(rightX, y);
                ctx.lineTo(rightX + dx, y - dy);
                ctx.lineTo(leftX + dx, y - dy);
                ctx.closePath();
                ctx.stroke();

                ctx.beginPath();
                ctx.moveTo(rightX, y);
                ctx.lineTo(rightX + dx, y - dy);
                ctx.lineTo(rightX + dx, bottomY - dy);
                ctx.lineTo(rightX, bottomY);
                ctx.closePath();
                ctx.stroke();

                // numeric label above top
                const value = ds.data[idx];
                if (typeof value !== 'undefined' && value !== null) {
                  const labelX = xCenter + dx / 2;
                  const labelY = y - dy - 2;
                  ctx.font = '700 28px "Helvetica Neue", Arial';
                  ctx.fillStyle = '#111827';
                  ctx.textAlign = 'center';
                  ctx.textBaseline = 'bottom';
                  ctx.fillText(String(value), labelX, labelY);
                }

                ctx.restore();
              });
            }
          }
        }
      ]
    });

    chart.update();

    await this.ensureUploadDir();
    const filename = `graph-3drect-${tableNumber}-${Date.now()}.png`;
    const filepath = path.join(UPLOAD_DIR, filename);
    const buffer = canvas.toBuffer('image/png');
    await fs.writeFile(filepath, buffer);

    try { chart.destroy(); } catch (e) { /* ignore */ }

    return filepath;
  }

  // fallback method with Y-axis fix
  async createHighQualityBarChart(tableData, tableNumber) {
    const width = 1600;
    const height = 1000;
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';

    if (!tableData || !Array.isArray(tableData.rows)) {
      throw new Error('Invalid tableData.rows');
    }

    const dataRows = tableData.rows.filter(r => String(r.category || '').toLowerCase() !== 'total');
    const labels = dataRows.map(r => String(r.category || ''));
    const trialData = dataRows.map(r => Number(r.trialGroup) || 0);
    const controlData = dataRows.map(r => Number(r.controlGroup) || 0);
    const totalData = dataRows.map(r => Number(r.total) || 0);

    const colors = {
      trial: { background: 'rgba(123,94,126,0.95)', border: 'rgba(90,60,90,1)' },
      control: { background: 'rgba(143,160,255,0.95)', border: 'rgba(100,120,210,1)' },
      total: { background: 'rgba(245,240,184,0.98)', border: 'rgba(220,200,100,1)' }
    };

    const chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: [
          { label: 'Trial Group', data: trialData, backgroundColor: colors.trial.background, borderColor: colors.trial.border, borderWidth: 2 },
          { label: 'Control Group', data: controlData, backgroundColor: colors.control.background, borderColor: colors.control.border, borderWidth: 2 },
          { label: 'Total', data: totalData, backgroundColor: colors.total.background, borderColor: colors.total.border, borderWidth: 2 }
        ]
      },
      options: {
        responsive: false,
        animation: false,
        maintainAspectRatio: false,
        plugins: {
          title: {
            display: true,
            text: `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''}`,
            font: { size: 28, weight: '700' },
            padding: 20
          },
          legend: { display: true, position: 'top' }
        },
        scales: {
          x: { ticks: { font: { size: 16, weight: '700' } }, grid: { display: false } },
          // FIXED: Changed Y-axis label here too
          y: { 
            beginAtZero: true, 
            ticks: { font: { size: 16, weight: '700' }, stepSize: 1 },
            title: { display: true, text: 'Number of patients (No. patients)', font: { size: 16, weight: '700' } }
          }
        }
      }
    });

    chart.update();

    await this.ensureUploadDir();
    const filename = `graph-hq-${tableNumber}-${Date.now()}.png`;
    const filepath = path.join(UPLOAD_DIR, filename);
    const buffer = canvas.toBuffer('image/png');
    await fs.writeFile(filepath, buffer);

    try { chart.destroy(); } catch (_) { }
    return filepath;
  }
}

module.exports = new GraphService();