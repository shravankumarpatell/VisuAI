const archiver = require('archiver');
const fs = require('fs');
const path = require('path');

class ZipService {
  async createZipArchive(graphs, originalFilename, type = 'default') {
    return new Promise((resolve, reject) => {
      try {
        const timestamp = Date.now();
        const outputPath = path.join('uploads', `graphs-${timestamp}.zip`);
        
        // Create a file to stream archive data to
        const output = fs.createWriteStream(outputPath);
        const archive = archiver('zip', {
          zlib: { level: 9 } // Sets the compression level
        });

        output.on('close', () => {
          console.log(`ZIP created: ${archive.pointer()} total bytes`);
          resolve(outputPath);
        });

        output.on('end', () => {
          console.log('Data has been drained');
        });

        archive.on('warning', (err) => {
          if (err.code === 'ENOENT') {
            console.warn('ZIP Warning:', err);
          } else {
            reject(err);
          }
        });

        archive.on('error', (err) => {
          reject(err);
        });

        // Pipe archive data to the file
        archive.pipe(output);

        // Add files to the archive
        graphs.forEach((graph, index) => {
          if (fs.existsSync(graph.imagePath)) {
            // Use a clean filename for the archive
            const archiveFilename = graph.filename || `graph-${index + 1}-${this.sanitizeFilename(graph.title)}.png`;
            archive.file(graph.imagePath, { name: archiveFilename });
            console.log(`Added to ZIP: ${archiveFilename}`);
          } else {
            console.warn(`File not found: ${graph.imagePath}`);
          }
        });

        // Add a readme file with metadata
        const readmeContent = this.generateReadmeContent(graphs, originalFilename, type);
        archive.append(readmeContent, { name: 'README.txt' });

        // Finalize the archive
        archive.finalize();

      } catch (error) {
        reject(error);
      }
    });
  }

  generateReadmeContent(graphs, originalFilename, type = 'default') {
    if (type === 'signs') {
      const content = `
Signs & Symptoms Assessment Analysis
===================================

Generated from: ${originalFilename}
Date: ${new Date().toLocaleDateString('en-US', { 
  weekday: 'long', 
  year: 'numeric', 
  month: 'long', 
  day: 'numeric' 
})}
Time: ${new Date().toLocaleTimeString()}

Total graphs: ${graphs.length}

Graph Details:
${graphs.map((graph, index) => 
  `${index + 1}. ${graph.filename || `signs-graph-${index + 1}.png`} - ${graph.title}`
).join('\n')}

Notes:
- All graphs are in PNG format with high quality
- Y-axis represents "Percentage (%)" with maximum scale of 100%
- Each graph shows percentage-based assessment data for different time periods:
  * 7th day (Green bars)
  * 14th day (Orange bars)
  * 21st day (Red bars)
  * 28th day (Purple bars)
- X-axis shows Signs & Symptoms categories
- Graphs are generated using 3D bar chart visualization
- Data represents Signs & Symptoms assessment over treatment periods

For questions or support, please contact the thesis helper system administrator.
`;
      return content.trim();
    }

    // Default content for regular tables
    const content = `
Thesis Statistical Analysis Graphs
==================================

Generated from: ${originalFilename}
Date: ${new Date().toLocaleDateString('en-US', { 
  weekday: 'long', 
  year: 'numeric', 
  month: 'long', 
  day: 'numeric' 
})}
Time: ${new Date().toLocaleTimeString()}

Total graphs: ${graphs.length}

Graph Details:
${graphs.map((graph, index) => 
  `${index + 1}. ${graph.filename || `graph-${index + 1}.png`} - ${graph.title}`
).join('\n')}

Notes:
- All graphs are in PNG format with high quality
- Y-axis represents "Number of patients (No. patients)"
- Each graph shows distribution data for Trial Group, Control Group, and Total
- Graphs are generated using 3D bar chart visualization

For questions or support, please contact the thesis helper system administrator.
`;
    return content.trim();
  }

  sanitizeFilename(text) {
    return text.replace(/[^a-zA-Z0-9-_]/g, '-').toLowerCase();
  }
}

module.exports = new ZipService();