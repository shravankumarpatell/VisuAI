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

  // NEW METHOD: Generate Signs & Symptoms graphs
  async generateSignsGraphs(tables) {
    if (!Array.isArray(tables)) throw new Error('tables must be an array');
    const graphs = [];
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      const imagePath = await this.createSignsAssessmentChart(table, i + 1);
      graphs.push({
        title: table.title || `Signs & Symptoms Analysis ${i + 1}`,
        tableNumber: i + 1,
        imagePath,
        filename: `signs-graph-${i + 1}-table-${table.tableNumber || i + 1}.png`
      });
    }
    return graphs;
  }

  // NEW METHOD: Generate Improvement Analysis Graph
  async generateImprovementGraph(improvementData) {
    if (!improvementData || !improvementData.categories || !improvementData.timePeriods) {
      throw new Error('Invalid improvement data structure');
    }

    console.log('Generating improvement analysis graph with data:', {
      categories: improvementData.categories.length,
      timePeriods: improvementData.timePeriods.length,
      hasGroupA: !!improvementData.groupA,
      hasGroupB: !!improvementData.groupB
    });

    const imagePath = await this.createImprovementAnalysisChart(improvementData);
    return imagePath;
  }

  // NEW METHOD: Create Improvement Analysis 3D Chart
// NEW METHOD: Create Improvement Analysis 3D Chart (Updated with Python logic)
// NEW METHOD: Create Improvement Analysis 3D Chart (Updated with Python logic)
// COMPLETELY REWRITTEN: Flexible createImprovementAnalysisChart method
async createImprovementAnalysisChart(improvementData) {
  // Dynamic sizing based on number of time periods and groups
  const timePeriodsCount = improvementData.timePeriods.length;
  const isSingleGroup = !improvementData.groupB || 
    Object.keys(improvementData.groupB).every(period => 
      Object.keys(improvementData.groupB[period]).every(category => 
        !improvementData.groupB[period][category] || improvementData.groupB[period][category].count === 0
      )
    );
  
  // Dynamic canvas sizing
  const baseWidth = 2400;
  const extraWidthPerPeriod = Math.max(0, (timePeriodsCount - 3) * 200); // Extra width for more than 3 periods
  const width = Math.min(4800, baseWidth + extraWidthPerPeriod); // Cap at 4800px
  const height = 1500;
  
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d');
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';

  const categories = improvementData.categories;
  const timePeriods = improvementData.timePeriods;
  const groupA = improvementData.groupA;
  const groupB = improvementData.groupB;

  console.log('Creating flexible improvement chart with data:', {
    categories: categories.length,
    timePeriods: timePeriods.length,
    timePeriodsArray: timePeriods,
    isSingleGroup: isSingleGroup,
    canvasWidth: width,
    canvasHeight: height
  });

  // Define chart area with dynamic margins
  const marginLeft = 160;
  const marginRight = isSingleGroup ? 
    Math.max(400, 300 + categories.length * 80) : // Dynamic margin based on category count
    Math.max(600, 400 + categories.length * 80);
  const marginTop = 120;
  const marginBottom = 220;
  
  const chartLeft = marginLeft;
  const chartRight = width - marginRight;
  const chartTop = marginTop;
  const chartBottom = height - marginBottom;

  // 3D effect parameters (scaled for canvas size)
  const dx = Math.max(6, Math.round((chartRight - chartLeft) * 0.008));
  const dy = Math.max(4, Math.round(dx * 0.45));
  const platformDx = Math.round(dx * 7);
  const platformDy = Math.round(dy * 6);

  // Calculate maximum value for scaling
  let maxTotal = 0;
  timePeriods.forEach(period => {
    categories.forEach(category => {
      const groupACount = groupA[period] && groupA[period][category] ? groupA[period][category].count : 0;
      maxTotal = Math.max(maxTotal, groupACount);
      
      if (!isSingleGroup && groupB[period] && groupB[period][category]) {
        const groupBCount = groupB[period][category].count || 0;
        maxTotal = Math.max(maxTotal, groupBCount);
      }
    });
  });
  
  const yMax = Math.max(5, Math.ceil(maxTotal / 5) * 5);

  // Clear canvas with white background
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  // Draw platform base
  ctx.save();
  ctx.fillStyle = '#9aa0a6';
  ctx.strokeStyle = '#4b4f52';
  ctx.lineWidth = 2.2;

  const leftX = chartLeft - dx - 6;
  const rightX = chartRight + 6;
  
  ctx.beginPath();
  ctx.moveTo(leftX, chartBottom);
  ctx.lineTo(rightX, chartBottom);
  ctx.lineTo(rightX + platformDx, chartBottom - platformDy);
  ctx.lineTo(leftX + platformDx, chartBottom - platformDy);
  ctx.closePath();
  ctx.fill();
  ctx.stroke();

  // Platform highlight
  ctx.fillStyle = 'rgba(255,255,255,0.06)';
  ctx.beginPath();
  ctx.moveTo(leftX + 6, chartBottom - 2);
  ctx.lineTo(rightX - 6, chartBottom - 2);
  ctx.lineTo(rightX - 6 + platformDx, chartBottom - platformDy - 2);
  ctx.lineTo(leftX + 6 + platformDx, chartBottom - platformDy - 2);
  ctx.closePath();
  ctx.fill();
  ctx.restore();

  // Draw Y-axis grid lines and labels
  const numSteps = 5;
  const fontSize = Math.max(20, Math.min(28, width / 100)); // Scale font size with canvas
  ctx.font = `600 ${fontSize}px "Helvetica Neue", Arial`;
  ctx.fillStyle = '#111827';
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';

  for (let i = 0; i <= numSteps; i++) {
    const yVal = (yMax / numSteps) * i;
    const yPixel = chartBottom - (yVal / yMax) * (chartBottom - chartTop);
    
    // Grid line
    ctx.beginPath();
    ctx.strokeStyle = 'rgba(0,0,0,0.08)';
    ctx.lineWidth = 1;
    ctx.moveTo(chartLeft, yPixel);
    ctx.lineTo(chartRight, yPixel);
    ctx.stroke();
    
    // Y-axis label
    ctx.fillText(Math.round(yVal).toString(), chartLeft - 12, yPixel);
  }

  // FLEXIBLE: Layout calculations for dynamic time period count and group count
  const barsPerTimepoint = categories.length;
  const timepointsPerGroup = timePeriods.length;
  const numGroups = isSingleGroup ? 1 : 2;
  
  // Dynamic spacing calculation
  const minGroupGap = isSingleGroup ? 0 : Math.max(40, (chartRight - chartLeft) * 0.05);
  const availableWidth = chartRight - chartLeft - minGroupGap;
  const groupSlot = availableWidth / numGroups;
  
  // Ensure minimum slot size even for many time periods
  const minTimepointSlot = 80;
  const timepointSlot = Math.max(minTimepointSlot, groupSlot / timepointsPerGroup);
  
  const catPadding = Math.max(4, timepointSlot * 0.04);
  const usableTimepointWidth = timepointSlot - 2 * catPadding;
  const barWidth = Math.max(8, (usableTimepointWidth / barsPerTimepoint) * 0.78);
  const barGap = barsPerTimepoint > 1 ? 
    Math.max(2, (usableTimepointWidth - barsPerTimepoint * barWidth) / (barsPerTimepoint - 1)) : 0;

  console.log('Layout calculations:', {
    numGroups,
    timepointsPerGroup,
    barsPerTimepoint,
    groupSlot,
    timepointSlot,
    barWidth,
    barGap
  });

  // Category colors (consistent)
  const categoryColors = [
    'rgb(110, 188, 96)',    // Cured - Green
    'rgb(245, 158, 11)',    // Marked improved - Orange  
    'rgb(239, 68, 68)',     // Moderate improved - Red
    'rgb(139, 92, 246)',    // Mild improved - Purple
    'rgb(239, 230, 178)'    // Not cured - Light yellow
  ];

  // FLEXIBLE: Calculate bar positions and draw bars for any number of groups/time periods
  const groups = isSingleGroup ? ['Group A'] : ['Group A', 'Group B'];
  const zeroVisualPx = 6;

  groups.forEach((groupName, groupIndex) => {
    const groupData = groupIndex === 0 ? groupA : groupB;
    const groupStart = chartLeft + groupIndex * (groupSlot + minGroupGap);
    
    timePeriods.forEach((timepoint, timepointIndex) => {
      const timepointStart = groupStart + timepointIndex * timepointSlot + catPadding;
      
      categories.forEach((category, categoryIndex) => {
        const barLeft = timepointStart + categoryIndex * (barWidth + barGap);
        const barRight = barLeft + barWidth;
        const barCenterX = barLeft + barWidth / 2;
        
        // Get data value
        const dataValue = groupData[timepoint] && groupData[timepoint][category] ? 
          groupData[timepoint][category].count : 0;
        
        const segmentHeight = dataValue > 0 ? 
          (dataValue / yMax) * (chartBottom - chartTop) : zeroVisualPx;
        
        const currentTop = chartBottom - segmentHeight;
        
        // Draw 3D bar
        const baseColor = categoryColors[categoryIndex % categoryColors.length];
        
        // Front face
        ctx.fillStyle = baseColor;
        ctx.fillRect(barLeft, currentTop, barWidth, segmentHeight);
        
        // Top face
        ctx.beginPath();
        ctx.moveTo(barLeft, currentTop);
        ctx.lineTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barLeft + dx, currentTop - dy);
        ctx.closePath();
        ctx.fill();
        
        // Right side face
        ctx.beginPath();
        ctx.moveTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barRight + dx, chartBottom - dy);
        ctx.lineTo(barRight, chartBottom);
        ctx.closePath();
        ctx.fill();
        
        // Outlines
        ctx.strokeStyle = '#111827';
        ctx.lineWidth = Math.max(1.5, dx * 0.25);
        
        // Front face outline
        ctx.strokeRect(barLeft, currentTop, barWidth, segmentHeight);
        
        // Top face outline
        ctx.beginPath();
        ctx.moveTo(barLeft, currentTop);
        ctx.lineTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barLeft + dx, currentTop - dy);
        ctx.closePath();
        ctx.stroke();
        
        // Right side outline
        ctx.beginPath();
        ctx.moveTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barRight + dx, chartBottom - dy);
        ctx.lineTo(barRight, chartBottom);
        ctx.closePath();
        ctx.stroke();
        
        // Value label on top of bar (scale font with bar size)
        if (dataValue > 0) {
          const labelFontSize = Math.max(16, Math.min(24, barWidth * 0.8));
          ctx.font = `700 ${labelFontSize}px "Helvetica Neue", Arial`;
          ctx.fillStyle = '#111827';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'bottom';
          const labelX = barCenterX + dx / 2;
          const labelY = currentTop - dy - 2;
          ctx.fillText(dataValue.toString(), labelX, labelY);
        }
      });
    });
  });

  // FLEXIBLE: X-axis timepoint labels with smart formatting
  const xLabelFontSize = Math.max(20, Math.min(32, Math.min(timepointSlot / 4, 32)));
  ctx.font = `700 ${xLabelFontSize}px "Helvetica Neue", Arial`;
  ctx.fillStyle = '#000000';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';

  groups.forEach((groupName, groupIndex) => {
    const groupStart = chartLeft + groupIndex * (groupSlot + minGroupGap);
    
    timePeriods.forEach((timepoint, timepointIndex) => {
      const timepointCenter = groupStart + timepointIndex * timepointSlot + timepointSlot / 2;
      
      // Smart formatting of time period labels
      let displayTimepoint = this.formatTimePeriodLabel(timepoint);
      
      ctx.fillText(displayTimepoint, timepointCenter, chartBottom + 18);
    });
  });

  // FLEXIBLE: Group labels
  const groupLabelFontSize = Math.max(24, Math.min(36, groupSlot / 8));
  ctx.font = `700 ${groupLabelFontSize}px "Helvetica Neue", Arial`;
  ctx.fillStyle = '#111827';
  
  if (isSingleGroup) {
    const groupACenterX = chartLeft + groupSlot / 2;
    ctx.fillText('Group A', groupACenterX, chartBottom + 68);
  } else {
    const groupACenterX = chartLeft + groupSlot / 2;
    const groupBCenterX = chartLeft + groupSlot + minGroupGap + groupSlot / 2;
    ctx.fillText('Group A', groupACenterX, chartBottom + 68);
    ctx.fillText('Group B', groupBCenterX, chartBottom + 68);
  }

  // FLEXIBLE: Dynamic title based on content
  const titleFontSize = Math.max(32, Math.min(52, width / 60));
  ctx.font = `700 ${titleFontSize}px "Helvetica Neue", Arial`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#111827';
  const titleX = chartLeft + (chartRight - chartLeft) / 2;
  
  const titleText = this.generateChartTitle(isSingleGroup, timePeriods.length, timePeriods);
  ctx.fillText(titleText, titleX, 40);

  // FLEXIBLE: Legend with dynamic positioning
  this.drawFlexibleLegend(ctx, chartRight, chartTop, categories, categoryColors, fontSize);

  // Y-axis title (rotated)
  ctx.save();
  ctx.translate(chartLeft - 120, chartTop + (chartBottom - chartTop) / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.font = `800 ${Math.max(20, Math.min(32, height / 50))}px "Helvetica Neue", Arial`;
  ctx.fillStyle = '#000000';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('Number of Patients', 0, 0);
  ctx.restore();

  // Save the chart
  await this.ensureUploadDir();
  const filename = `improvement-analysis-${isSingleGroup ? 'single-group' : 'comparison'}-${timePeriods.length}periods-${Date.now()}.png`;
  const filepath = path.join(UPLOAD_DIR, filename);
  const buffer = canvas.toBuffer('image/png');
  await fs.writeFile(filepath, buffer);

  console.log(`Flexible improvement analysis graph created (${isSingleGroup ? 'single group' : 'dual group'}, ${timePeriods.length} periods):`, filepath);
  return filepath;
}

// NEW: Smart time period label formatting
formatTimePeriodLabel(timepoint) {
  // Handle different time period formats
  if (/^\d+(st|nd|rd|th)$/.test(timepoint)) {
    // Ordinal numbers: "7th" -> "7th day"
    return `${timepoint} day`;
  } else if (/^[A-Z]{1,3}$/.test(timepoint)) {
    // Letter codes: keep as is
    return timepoint;
  } else if (/baseline|initial|pre/i.test(timepoint)) {
    return 'Baseline';
  } else if (/follow|final|post/i.test(timepoint)) {
    return 'Follow-up';
  } else if (/^\d+$/.test(timepoint)) {
    // Plain numbers: add "day"
    return `${timepoint} day`;
  } else {
    // Return as is for any other format
    return timepoint;
  }
}

// NEW: Generate contextual chart title
generateChartTitle(isSingleGroup, periodCount, periods) {
  let baseTitle = 'Improvement Analysis';
  
  if (isSingleGroup) {
    baseTitle += ' - Group A';
  } else {
    baseTitle += ' - Group A vs Group B Comparison';
  }
  
  // Add period context if meaningful
  if (periodCount === 1) {
    baseTitle += ` (Single Period)`;
  } else if (periodCount <= 3) {
    baseTitle += ` (${periodCount} Periods)`;
  }
  
  return baseTitle;
}

// NEW: Flexible legend drawing
drawFlexibleLegend(ctx, legendX, legendY, categories, categoryColors, baseFontSize) {
  const boxSize = Math.max(24, Math.min(40, baseFontSize * 1.2));
  const fontSize = Math.max(20, Math.min(32, baseFontSize * 0.9));
  const baseLineHeight = Math.max(50, fontSize * 2.2);
  
  ctx.font = `600 ${fontSize}px "Helvetica Neue", Arial`;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  
  let currentY = legendY + 20;
  const maxTextWidth = 300; // Maximum width for legend text
  
  categories.forEach((category, index) => {
    // Legend color box
    ctx.fillStyle = categoryColors[index % categoryColors.length];
    ctx.fillRect(legendX + 20, currentY, boxSize, boxSize);
    ctx.strokeStyle = '#111827';
    ctx.lineWidth = 1;
    ctx.strokeRect(legendX + 20, currentY, boxSize, boxSize);
    
    // Legend text with smart wrapping
    ctx.fillStyle = '#111827';
    const textX = legendX + boxSize + 40;
    
    let displayText = category;
    let line1 = '';
    let line2 = '';
    let needsWrapping = false;
    
    // Check if text needs wrapping
    const textMetrics = ctx.measureText(displayText);
    if (textMetrics.width > maxTextWidth) {
      needsWrapping = true;
      
      // Smart text breaking for medical categories
      if (category.includes('(') && category.includes(')')) {
        const parenIndex = category.indexOf('(');
        line1 = category.substring(0, parenIndex).trim();
        line2 = category.substring(parenIndex).trim();
      } else if (category.length > 15) {
        // Break at logical points
        const words = category.split(' ');
        const midPoint = Math.ceil(words.length / 2);
        line1 = words.slice(0, midPoint).join(' ');
        line2 = words.slice(midPoint).join(' ');
      } else {
        // Character-based break
        const midPoint = Math.ceil(category.length / 2);
        line1 = category.substring(0, midPoint);
        line2 = category.substring(midPoint);
      }
    }
    
    if (needsWrapping && line1 && line2) {
      // Draw two lines
      const boxCenterY = currentY + boxSize / 2;
      ctx.fillText(line1, textX, boxCenterY - fontSize * 0.6);
      ctx.fillText(line2, textX, boxCenterY + fontSize * 0.6);
      currentY += baseLineHeight * 1.2;
    } else {
      // Draw single line
      const boxCenterY = currentY + boxSize / 2;
      ctx.fillText(displayText, textX, boxCenterY);
      currentY += baseLineHeight;
    }
  });
}

  // NEW METHOD: Create Signs & Symptoms Assessment Chart
async createImprovementAnalysisChart(improvementData) {
  const width = 3200;
  const height = 1500;
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext('2d');
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';

  const categories = improvementData.categories;
  const timePeriods = improvementData.timePeriods; // This should include ALL periods including 21st
  const groupA = improvementData.groupA;
  const groupB = improvementData.groupB;

  // FIXED: Detect if this is a single group analysis
  const isSingleGroup = !groupB || Object.keys(groupB).every(period => 
    Object.keys(groupB[period]).every(category => 
      !groupB[period][category] || groupB[period][category].count === 0
    )
  );

  console.log('Creating improvement chart with data:', {
    categories: categories.length,
    timePeriods: timePeriods.length,
    timePeriodsList: timePeriods,
    isSingleGroup: isSingleGroup,
    groupAData: Object.keys(groupA),
    groupBData: groupB ? Object.keys(groupB) : 'none'
  });

  // Define chart area
  const marginLeft = 160;
  const marginRight = isSingleGroup ? 400 : 600; // Smaller margin for single group
  const marginTop = 120;
  const marginBottom = 220;
  
  const chartLeft = marginLeft;
  const chartRight = width - marginRight;
  const chartTop = marginTop;
  const chartBottom = height - marginBottom;

  // 3D effect parameters
  const dx = Math.max(8, Math.round((chartRight - chartLeft) * 0.01));
  const dy = Math.max(6, Math.round(dx * 0.45));
  const platformDx = Math.round(dx * 7);
  const platformDy = Math.round(dy * 6);

  // Calculate maximum value for scaling - check ALL time periods
  let maxTotal = 0;
  timePeriods.forEach(period => {
    categories.forEach(category => {
      const groupACount = groupA[period] && groupA[period][category] ? groupA[period][category].count : 0;
      maxTotal = Math.max(maxTotal, groupACount);
      
      // Only check Group B if it's not a single group
      if (!isSingleGroup && groupB[period] && groupB[period][category]) {
        const groupBCount = groupB[period][category].count || 0;
        maxTotal = Math.max(maxTotal, groupBCount);
      }
    });
  });
  
  const yMax = Math.max(5, Math.ceil(maxTotal / 5) * 5);

  // Clear canvas with white background
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  // Draw platform base
  ctx.save();
  ctx.fillStyle = '#9aa0a6';
  ctx.strokeStyle = '#4b4f52';
  ctx.lineWidth = 2.2;

  const leftX = chartLeft - dx - 6;
  const rightX = chartRight + 6;
  
  ctx.beginPath();
  ctx.moveTo(leftX, chartBottom);
  ctx.lineTo(rightX, chartBottom);
  ctx.lineTo(rightX + platformDx, chartBottom - platformDy);
  ctx.lineTo(leftX + platformDx, chartBottom - platformDy);
  ctx.closePath();
  ctx.fill();
  ctx.stroke();

  // Platform highlight
  ctx.fillStyle = 'rgba(255,255,255,0.06)';
  ctx.beginPath();
  ctx.moveTo(leftX + 6, chartBottom - 2);
  ctx.lineTo(rightX - 6, chartBottom - 2);
  ctx.lineTo(rightX - 6 + platformDx, chartBottom - platformDy - 2);
  ctx.lineTo(leftX + 6 + platformDx, chartBottom - platformDy - 2);
  ctx.closePath();
  ctx.fill();
  ctx.restore();

  // Draw Y-axis grid lines and labels
  const numSteps = 5;
  ctx.font = '600 24px "Helvetica Neue", Arial';
  ctx.fillStyle = '#111827';
  ctx.textAlign = 'right';
  ctx.textBaseline = 'middle';

  for (let i = 0; i <= numSteps; i++) {
    const yVal = (yMax / numSteps) * i;
    const yPixel = chartBottom - (yVal / yMax) * (chartBottom - chartTop);
    
    // Grid line
    ctx.beginPath();
    ctx.strokeStyle = 'rgba(0,0,0,0.08)';
    ctx.lineWidth = 1;
    ctx.moveTo(chartLeft, yPixel);
    ctx.lineTo(chartRight, yPixel);
    ctx.stroke();
    
    // Y-axis label
    ctx.fillText(Math.round(yVal).toString(), chartLeft - 12, yPixel);
  }

  // FIXED: Layout calculations for dynamic group count
  const barsPerTimepoint = categories.length; // 5 categories
  const timepointsPerGroup = timePeriods.length; // ALL timepoints (4 or 5)
  const numGroups = isSingleGroup ? 1 : 2;
  const groupGap = isSingleGroup ? 0 : 0.08 * (chartRight - chartLeft);
  const availableWidth = chartRight - chartLeft - groupGap;
  const groupSlot = availableWidth / numGroups;
  const timepointSlot = groupSlot / timepointsPerGroup;
  const catPadding = 0.06 * timepointSlot;
  const usableTimepointWidth = timepointSlot - 2 * catPadding;
  const barWidth = (usableTimepointWidth / barsPerTimepoint) * 0.78;
  const barGap = barsPerTimepoint > 1 ? 
    (usableTimepointWidth - barsPerTimepoint * barWidth) / (barsPerTimepoint - 1) : 0;

  // Category colors
  const categoryColors = [
    'rgb(110, 188, 96)',    // Cured - Green
    'rgb(245, 158, 11)',    // Marked improved - Orange  
    'rgb(239, 68, 68)',     // Moderate improved - Red
    'rgb(139, 92, 246)',    // Mild improved - Purple
    'rgb(239, 230, 178)'    // Not cured - Light yellow
  ];

  // FIXED: Calculate bar positions and draw bars for dynamic groups
  const groups = isSingleGroup ? ['Group A'] : ['Group A', 'Group B'];
  const zeroVisualPx = 6;

  groups.forEach((groupName, groupIndex) => {
    const groupData = groupIndex === 0 ? groupA : groupB;
    const groupStart = chartLeft + groupIndex * (groupSlot + groupGap);
    
    timePeriods.forEach((timepoint, timepointIndex) => {
      const timepointStart = groupStart + timepointIndex * timepointSlot + catPadding;
      
      categories.forEach((category, categoryIndex) => {
        const barLeft = timepointStart + categoryIndex * (barWidth + barGap);
        const barRight = barLeft + barWidth;
        const barCenterX = barLeft + barWidth / 2;
        
        // Get data value
        const dataValue = groupData[timepoint] && groupData[timepoint][category] ? 
          groupData[timepoint][category].count : 0;
        
        const segmentHeight = dataValue > 0 ? 
          (dataValue / yMax) * (chartBottom - chartTop) : zeroVisualPx;
        
        const currentTop = chartBottom - segmentHeight;
        
        // Draw 3D bar
        const baseColor = categoryColors[categoryIndex];
        
        // Front face
        ctx.fillStyle = baseColor;
        ctx.fillRect(barLeft, currentTop, barWidth, segmentHeight);
        
        // Top face
        ctx.beginPath();
        ctx.moveTo(barLeft, currentTop);
        ctx.lineTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barLeft + dx, currentTop - dy);
        ctx.closePath();
        ctx.fill();
        
        // Right side face
        ctx.beginPath();
        ctx.moveTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barRight + dx, chartBottom - dy);
        ctx.lineTo(barRight, chartBottom);
        ctx.closePath();
        ctx.fill();
        
        // Outlines
        ctx.strokeStyle = '#111827';
        ctx.lineWidth = 2;
        
        // Front face outline
        ctx.strokeRect(barLeft, currentTop, barWidth, segmentHeight);
        
        // Top face outline
        ctx.beginPath();
        ctx.moveTo(barLeft, currentTop);
        ctx.lineTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barLeft + dx, currentTop - dy);
        ctx.closePath();
        ctx.stroke();
        
        // Right side outline
        ctx.beginPath();
        ctx.moveTo(barRight, currentTop);
        ctx.lineTo(barRight + dx, currentTop - dy);
        ctx.lineTo(barRight + dx, chartBottom - dy);
        ctx.lineTo(barRight, chartBottom);
        ctx.closePath();
        ctx.stroke();
        
        // Value label on top of bar
        if (dataValue > 0) {
          ctx.font = '700 22px "Helvetica Neue", Arial';
          ctx.fillStyle = '#111827';
          ctx.textAlign = 'center';
          ctx.textBaseline = 'bottom';
          const labelX = barCenterX + dx / 2;
          const labelY = currentTop - dy - 6;
          ctx.fillText(dataValue.toString(), labelX, labelY);
        }
      });
    });
  });

  // FIXED: X-axis timepoint labels for all periods
  ctx.font = '700 30px "Helvetica Neue", Arial';
  ctx.fillStyle = '#000000';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';

  groups.forEach((groupName, groupIndex) => {
    const groupStart = chartLeft + groupIndex * (groupSlot + groupGap);
    
    timePeriods.forEach((timepoint, timepointIndex) => {
      const timepointCenter = groupStart + timepointIndex * timepointSlot + timepointSlot / 2;
      // FIXED: Format timepoint labels properly
      const displayTimepoint = timepoint.replace('th', 'th day').replace('st', 'st day').replace('nd', 'nd day').replace('rd', 'rd day');
      ctx.fillText(displayTimepoint, timepointCenter, chartBottom + 18);
    });
  });

  // FIXED: Group labels for dynamic groups
  ctx.font = '700 30px "Helvetica Neue", Arial';
  ctx.fillStyle = '#111827';
  
  if (isSingleGroup) {
    const groupACenterX = chartLeft + groupSlot / 2;
    ctx.fillText('Group A', groupACenterX, chartBottom + 68);
  } else {
    const groupACenterX = chartLeft + groupSlot / 2;
    const groupBCenterX = chartLeft + groupSlot + groupGap + groupSlot / 2;
    ctx.fillText('Group A', groupACenterX, chartBottom + 68);
    ctx.fillText('Group B', groupBCenterX, chartBottom + 68);
  }

  // FIXED: Dynamic title
  ctx.font = '700 48px "Helvetica Neue", Arial';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillStyle = '#111827';
  const titleX = chartLeft + (chartRight - chartLeft) / 2;
  const titleText = isSingleGroup ? 
    'Improvement Analysis - Group A' : 
    'Improvement Analysis - Group A vs Group B Comparison';
  ctx.fillText(titleText, titleX, 40);

  // Legend
const legendX = chartRight + 60;
const legendY = chartTop + 40;
const boxSize = 32;
const baseLineHeight = 75; // Increased base line height to accommodate wrapped text

ctx.font = '600 26px "Helvetica Neue", Arial';
ctx.textAlign = 'left';
ctx.textBaseline = 'middle';

let currentY = legendY;

categories.forEach((category, index) => {
  // Legend color box
  ctx.fillStyle = categoryColors[index];
  ctx.fillRect(legendX, currentY, boxSize, boxSize);
  ctx.strokeStyle = '#111827';
  ctx.lineWidth = 1;
  ctx.strokeRect(legendX, currentY, boxSize, boxSize);
  
  // Legend text with improved wrapping
  ctx.fillStyle = '#111827';
  const textX = legendX + boxSize + 20;
  const maxTextWidth = width - textX - 30; // Leave 30px margin from right edge
  
  // Smart text processing for better readability
  let displayText = category;
  let line1 = '';
  let line2 = '';
  let needsWrapping = false;
  
  // Check if text is too wide
  const textMetrics = ctx.measureText(displayText);
  if (textMetrics.width > maxTextWidth) {
    needsWrapping = true;
    
    // Handle specific patterns for medical/improvement categories
    if (category.includes('improved(') || category.includes('Improved(')) {
      // For "Marked improved(75-100%)" -> "Marked improved" + "(75-100%)"
      const parenIndex = category.indexOf('(');
      if (parenIndex > 0) {
        line1 = category.substring(0, parenIndex).trim();
        line2 = category.substring(parenIndex).trim();
      }
    } else if (category.includes('cured(') || category.includes('Cured(')) {
      // For "Not cured(<25%)" -> "Not cured" + "(<25%)"
      const parenIndex = category.indexOf('(');
      if (parenIndex > 0) {
        line1 = category.substring(0, parenIndex).trim();
        line2 = category.substring(parenIndex).trim();
      }
    } else if (category.length > 15) {
      // For other long text, find the best breaking point
      const words = category.split(' ');
      let testLine = '';
      let breakIndex = Math.ceil(words.length / 2); // Start from middle
      
      // Find optimal break point
      for (let i = 1; i < words.length; i++) {
        const testText = words.slice(0, i).join(' ');
        if (ctx.measureText(testText).width > maxTextWidth * 0.6) {
          breakIndex = Math.max(1, i - 1);
          break;
        }
      }
      
      line1 = words.slice(0, breakIndex).join(' ');
      line2 = words.slice(breakIndex).join(' ');
    }
    
    // Fallback if we couldn't break it properly
    if (!line1 && !line2) {
      line1 = category.substring(0, Math.ceil(category.length / 2));
      line2 = category.substring(Math.ceil(category.length / 2));
    }
  }
  
  if (needsWrapping && line1 && line2) {
    // Draw two lines
    const boxCenterY = currentY + boxSize / 2;
    ctx.fillText(line1, textX, boxCenterY - 12);
    ctx.fillText(line2, textX, boxCenterY + 12);
    currentY += baseLineHeight; // Use larger spacing for wrapped text
  } else {
    // Draw single line
    const boxCenterY = currentY + boxSize / 2;
    ctx.fillText(displayText, textX, boxCenterY);
    currentY += Math.max(baseLineHeight - 10, 65); // Slightly smaller spacing for single lines
  }
});

  // Y-axis title (rotated)
  ctx.save();
  ctx.translate(chartLeft - 120, chartTop + (chartBottom - chartTop) / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.font = '800 26px "Helvetica Neue", Arial';
  ctx.fillStyle = '#000000';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('Number of Patients', 0, 0);
  ctx.restore();

  // Save the chart
  await this.ensureUploadDir();
  const filename = `improvement-analysis-${isSingleGroup ? 'single-group' : 'comparison'}-${Date.now()}.png`;
  const filepath = path.join(UPLOAD_DIR, filename);
  const buffer = canvas.toBuffer('image/png');
  await fs.writeFile(filepath, buffer);

  console.log(`Improvement analysis graph created (${isSingleGroup ? 'single group' : 'dual group'}):`, filepath);
  return filepath;
}


// NEW METHOD: Create Signs & Symptoms Assessment Chart
async createSignsAssessmentChart(tableData, tableNumber) {
    const width = 2400;
    const height = 1500;
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';

    if (!tableData || !Array.isArray(tableData.signs) || !tableData.data) {
      throw new Error('Invalid Signs tableData structure');
    }

    const labels = tableData.signs; // ["Vedana", "Varna", "Sraava", etc.]
    const assessmentPeriods = tableData.assessments; // ["7 day", "14 day", "21 day", "28 day"]
    
    // Create datasets for each assessment period
    const datasets = [];
    const colors = ['#4ade80', '#f59e0b', '#ef4444', '#8b5cf6']; // Green, Orange, Red, Purple
    
    assessmentPeriods.forEach((period, index) => {
      const data = labels.map(sign => tableData.data[sign] ? tableData.data[sign][period] || 0 : 0);
      
      datasets.push({
        label: `${period}`,
        data: data,
        backgroundColor: colors[index % colors.length],
        borderColor: '#000000ff',
        borderWidth: 2,
        barPercentage: 0.6,
        categoryPercentage: 0.8
      });
    });

    const chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets
      },
      options: {
        responsive: false,
        animation: false,
        maintainAspectRatio: false,
        layout: { padding: { left: 120, right: 260, top: 90, bottom: 140 } },
        plugins: {
          title: {
            display: true,
            text: `Graph no ${tableNumber} - Signs & Symptoms Assessment Analysis`,
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
          tooltip: { 
            enabled: true,
            callbacks: {
              label: function(context) {
                return `${context.dataset.label}: ${context.parsed.y}%`;
              }
            }
          }
        },
        scales: {
          x: {
            ticks: { 
              font: { size: 20, weight: '600' }, 
              color: '#000000ff',
              maxRotation: 45,
              minRotation: 0
            },
            grid: { display: false, drawBorder: true, borderWidth: 4, borderColor: '#111827' },
            title: { 
              display: true, 
              text: 'Signs & Symptoms', 
              font: { size: 30, weight: '800' }, 
              color: '#000000ff' 
            }
          },
          y: {
            beginAtZero: true,
            max: 100,
            ticks: { 
              font: { size: 24, weight: '600' }, 
              stepSize: 10, 
              color: '#000000ff',
              callback: function(value) {
                return value + '%';
              }
            },
            grid: { 
              display: true, 
              color: 'rgba(0,0,0,0.06)', 
              lineWidth: 1, 
              drawBorder: true, 
              borderWidth: 4, 
              borderColor: '#111827' 
            },
            title: { 
              display: true, 
              text: 'Percentage (%)', 
              font: { size: 30, weight: '800' }, 
              color: '#000000ff' 
            }
          }
        },
        elements: { bar: { borderWidth: 5 } }
      },
      plugins: [
        {
          id: 'signs-3d-rect-platform',
          beforeDatasetsDraw: chartInstance => {
            const ctx = chartInstance.ctx;
            const chartArea = chartInstance.chartArea;

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

            ctx.restore();
          },

          afterDatasetsDraw: chartInstance => {
            const ctx = chartInstance.ctx;
            const chartArea = chartInstance.chartArea;
            const datasets = chartInstance.data.datasets;

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

                // Outlines
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

                // Numeric label above top
                const value = ds.data[idx];
                if (typeof value !== 'undefined' && value !== null) {
                  const labelX = leftX + w/2 + dx / 2;
                  const labelY = y - dy - 2;
                  ctx.font = '700 24px "Helvetica Neue", Arial';
                  ctx.fillStyle = '#111827';
                  ctx.textAlign = 'center';
                  ctx.textBaseline = 'bottom';
                  ctx.fillText(String(value) + '%', labelX, labelY);
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
    const filename = `signs-3drect-${tableNumber}-${Date.now()}.png`;
    const filepath = path.join(UPLOAD_DIR, filename);
    const buffer = canvas.toBuffer('image/png');
    await fs.writeFile(filepath, buffer);

    try { chart.destroy(); } catch (e) { /* ignore */ }

    return filepath;
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

    // Check if this is a single-sheet table (no control group data)
    const isSingleSheet = tableData.isPaired === false || controlData.every(val => val === 0);
    
    const palette = Object.assign({
      trial: '#9fb4ffff',
      control: '#8e4b6eff',
      total: '#efe6b2ff'
    }, tableData.colors || {});

    const xAxisTitle = tableData.xLabel || tableData.xTitle || (tableData.title && tableData.title.split('-').pop().trim()) || 'Category';
    const yAxisTitle = tableData.yLabel || tableData.yTitle || 'No. patients';

    // Create datasets - exclude control group if it's a single sheet table
    const datasets = [
      { 
        label: isSingleSheet ? (tableData.sourceInfo?.trialGroup || 'Group Data') : 'Trial Group', 
        data: trialData, 
        backgroundColor: palette.trial, 
        borderColor: '#000000ff', 
        borderWidth: 2, 
        barPercentage: 0.5, 
        categoryPercentage: 0.6 
      }
    ];

    // Only add control group and total if we have paired data
    if (!isSingleSheet) {
      datasets.push(
        { 
          label: 'Control Group', 
          data: controlData, 
          backgroundColor: palette.control, 
          borderColor: '#000000ff', 
          borderWidth: 2, 
          barPercentage: 0.5, 
          categoryPercentage: 0.6 
        },
        { 
          label: 'Total', 
          data: totalData, 
          backgroundColor: palette.total, 
          borderColor: '#000000ff', 
          borderWidth: 2, 
          barPercentage: 0.5, 
          categoryPercentage: 0.6 
        }
      );
    }

    // Adjust title for single sheet
    const graphTitle = isSingleSheet ? 
      `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''}` :
      `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''}`;

    const chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets
      },
      options: {
        responsive: false,
        animation: false,
        maintainAspectRatio: false,
        layout: { padding: { left: 120, right: 260, top: 90, bottom: 140 } },
        plugins: {
          title: {
            display: true,
            text: graphTitle,
            color: '#111827',
            font: { size: isSingleSheet ? 32 : 36, weight: '700', family: "'Helvetica Neue', Arial" },
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
            grid: { display: true, color: 'rgba(0,0,0,0.06)', lineWidth: 1, drawBorder: true, borderWidth: 4, borderColor: '#111827' },
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

  // fallback method with Y-axis fix and single-sheet support
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

    // Check if this is a single-sheet table
    const isSingleSheet = tableData.isPaired === false || controlData.every(val => val === 0);

    const colors = {
      trial: { background: 'rgba(123,94,126,0.95)', border: 'rgba(90,60,90,1)' },
      control: { background: 'rgba(143,160,255,0.95)', border: 'rgba(100,120,210,1)' },
      total: { background: 'rgba(245,240,184,0.98)', border: 'rgba(220,200,100,1)' }
    };

    // Create datasets - exclude control group if it's a single sheet table
    const datasets = [
      { 
        label: isSingleSheet ? (tableData.sourceInfo?.trialGroup || 'Data Group') : 'Trial Group', 
        data: trialData, 
        backgroundColor: colors.trial.background, 
        borderColor: colors.trial.border, 
        borderWidth: 2 
      }
    ];

    if (!isSingleSheet) {
      datasets.push(
        { label: 'Control Group', data: controlData, backgroundColor: colors.control.background, borderColor: colors.control.border, borderWidth: 2 },
        { label: 'Total', data: totalData, backgroundColor: colors.total.background, borderColor: colors.total.border, borderWidth: 2 }
      );
    }

    const graphTitle = isSingleSheet ? 
      `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''} (${tableData.sourceInfo?.trialGroup || 'Single Dataset'})` :
      `Graph no ${tableNumber} - Distribution of subjects based on ${tableData.title || ''}`;

    const chart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets
      },
      options: {
        responsive: false,
        animation: false,
        maintainAspectRatio: false,
        plugins: {
          title: {
            display: true,
            text: graphTitle,
            font: { size: isSingleSheet ? 24 : 28, weight: '700' },
            padding: 20
          },
          legend: { display: true, position: 'top' }
        },
        scales: {
          x: { ticks: { font: { size: 16, weight: '700' } }, grid: { display: false } },
          y: {
            beginAtZero: true,
            ticks: { font: { size: 16, weight: '700' }, stepSize: 1 },
            title: { display: true, text: 'No. patients', font: { size: 16, weight: '700' } }
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