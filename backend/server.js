const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
require('dotenv').config();

// Import services
const excelService = require('./services/excelService');
const wordService = require('./services/wordService');
const wordParserService = require('./services/wordParserService');
const signsParserService = require('./services/signsParserService');
const masterChartService = require('./services/masterChartService');
if (!masterChartService || typeof masterChartService.processMasterChart !== 'function') {
  console.error('ERROR: services/masterChartService does not export processMasterChart(). Check services/masterChartService.js export.');
  throw new Error('masterChartService.processMasterChart is not available');
}
const graphService = require('./services/graphService');
const pdfService = require('./services/pdfService');
const zipService = require('./services/zipService');

// Import utilities
const ValidationUtils = require('./utils/validation');
const CleanupUtility = require('./utils/cleanup');
const { errorHandler, asyncHandler, notFound, AppError } = require('./middleware/errorHandler');

const app = express();
const PORT = process.env.PORT || 5000;

// Initialize cleanup utility
const cleanup = new CleanupUtility(
  process.env.UPLOAD_DIR || 'uploads',
  parseInt(process.env.FILE_RETENTION_HOURS) || 24
);

// Ensure upload directory exists
cleanup.ensureUploadDirectory();

// Start periodic cleanup
cleanup.startPeriodicCleanup(6);

// Middleware
app.use(cors({
  origin: process.env.ALLOWED_ORIGINS ? process.env.ALLOWED_ORIGINS.split(',') : '*',
  credentials: true
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: async (req, file, cb) => {
    const uploadDir = process.env.UPLOAD_DIR || 'uploads';
    try {
      await fs.mkdir(uploadDir, { recursive: true });
    } catch (error) {
      console.error('Error creating upload directory:', error);
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const sanitizedFilename = ValidationUtils.sanitizeFilename(file.originalname);
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    const ext = path.extname(sanitizedFilename);
    const name = path.basename(sanitizedFilename, ext);
    cb(null, `${name}-${uniqueSuffix}${ext}`);
  }
});

const upload = multer({ 
  storage: storage,
  limits: {
    fileSize: (parseInt(process.env.MAX_FILE_SIZE_MB) || 10) * 1024 * 1024
  },
  fileFilter: (req, file, cb) => {
    const allowedTypes = {
      'excel-to-word': ['.xlsx', '.xls'],
      'word-to-pdf': ['.docx', '.doc'],
      'signs-analysis': ['.docx', '.doc'],
      'master-chart-analysis': ['.xlsx', '.xls'],
      'improvement-analysis': ['.docx', '.doc'] // NEW
    };
    
    let route = 'word-to-pdf'; // default
    if (req.path.includes('excel-to-word')) route = 'excel-to-word';
    else if (req.path.includes('signs-analysis')) route = 'signs-analysis';
    else if (req.path.includes('master-chart-analysis')) route = 'master-chart-analysis';
    else if (req.path.includes('improvement-analysis')) route = 'improvement-analysis'; // NEW
    
    const ext = path.extname(file.originalname).toLowerCase();
    
    if (!ValidationUtils.validateFileExtension(file.originalname, allowedTypes[route])) {
      cb(new AppError(`Invalid file type. Allowed types: ${allowedTypes[route].join(', ')}`, 400));
    } else {
      cb(null, true);
    }
  }
});

// NEW ROUTE: Improvement Analysis
app.post('/api/improvement-analysis', upload.single('file'), asyncHandler(async (req, res) => {
  if (!req.file) {
    throw new AppError('No file uploaded', 400);
  }

  console.log('Processing Improvement Analysis Word file:', req.file.filename);

  try {
    // Validate file size
    if (!ValidationUtils.validateFileSize(req.file.size, parseInt(process.env.MAX_FILE_SIZE_MB) || 10)) {
      throw new AppError('File size exceeds limit', 400);
    }

    // Parse Word document for improvement data
    const improvementData = await wordParserService.parseImprovementDocument(req.file.path);
    
    if (!improvementData) {
      throw new AppError('No improvement data found in the document. Please ensure your document contains a table with improvement categories and time period data.', 422);
    }

    // UPDATED: Use flexible validation that accepts any time period format
    const validation = ValidationUtils.validateImprovementDocument(improvementData);
    if (!validation.valid) {
      console.error('Improvement document validation failed:', validation.errors);
      throw new AppError(`Invalid improvement document format: ${validation.errors[0]}`, 422);
    }
    
    console.log(`Validation passed for improvement data with ${improvementData.timePeriods.length} time periods: [${improvementData.timePeriods.join(', ')}]`);
    
    // Generate improvement analysis graph (single PNG file)
    const graphPath = await graphService.generateImprovementGraph(improvementData);
    
    if (!graphPath) {
      throw new AppError('Failed to generate improvement analysis graph', 500);
    }
    
    console.log('Improvement analysis graph created:', graphPath);
    
    // Send PNG file
    res.download(graphPath, 'improvement-analysis-graph.png', async (err) => {
      if (err) {
        console.error('Error sending improvement graph file:', err);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error downloading improvement graph file' });
        }
      }
      
      // Cleanup
      await cleanup.cleanupFiles([req.file.path, graphPath]);
    });

  } catch (error) {
    console.error('Improvement analysis error:', error);
    
    // Cleanup on error
    await cleanup.cleanupFiles([req.file.path]);
    
    // Enhance error messages for common issues
    let errorMessage = error.message;
    
    if (error.message.includes('No improvement data found')) {
      errorMessage = 'Could not find improvement data in the document. Please ensure your document contains a table with:' +
        '\n- Improvement categories (like "Cured", "Marked improved", etc.)' +
        '\n- Time period columns (can be any format: "7th day", "AT", "AF", "1st", "3rd", etc.)' +
        '\n- Group A and/or Group B data';
    } else if (error.message.includes('time periods')) {
      errorMessage = `Time period parsing issue: ${error.message}. Supported formats include: 7th, 14th, 21st, AT, AF, 1st, 3rd, 5th, 19th, etc.`;
    }
    
    throw new AppError(errorMessage, error.statusCode || 500);
  }
}));

// EXISTING ROUTE: Master Chart Analysis
app.post('/api/master-chart-analysis', upload.single('file'), asyncHandler(async (req, res) => {
  if (!req.file) {
    throw new AppError('No file uploaded', 400);
  }

  console.log('Processing Master Chart Excel file:', req.file.filename);

  try {
    // Validate file size
    if (!ValidationUtils.validateFileSize(req.file.size, parseInt(process.env.MAX_FILE_SIZE_MB) || 10)) {
      throw new AppError('File size exceeds limit', 400);
    }

    // Process Master Chart Excel file
    const masterChartData = await masterChartService.processMasterChart(req.file.path);
    
    // Validate master chart data (best-effort: if validation fails we proceed but log warnings)
    let validation = { valid: true, errors: [] };
    try {
      if (ValidationUtils && typeof ValidationUtils.validateMasterChartData === 'function') {
        validation = ValidationUtils.validateMasterChartData(masterChartData);
      }
    } catch (valErr) {
      console.warn('Master chart validation threw an exception, continuing with best-effort processing:', valErr && valErr.message ? valErr.message : valErr);
      validation = { valid: false, errors: [String(valErr)] };
    }
    
    if (!validation.valid) {
      // Log validation errors but proceed – this allows generating output for slightly non-standard master charts.
      console.warn('Master chart validation failed with errors:', validation.errors);
    }
    
    // Generate Word document with complete statistical tables
    const wordFilePath = await wordService.generateMasterChartWordDocument(masterChartData, req.file.filename);
    
    // Send file
    res.download(wordFilePath, 'master-chart-analysis.docx', async (err) => {
      if (err) {
        console.error('Error sending file:', err);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error downloading file' });
        }
      }
      
      // Cleanup
      await cleanup.cleanupFiles([req.file.path, wordFilePath]);
    });

  } catch (error) {
    // Cleanup on error
    await cleanup.cleanupFiles([req.file.path]);
    throw error;
  }
}));

// Existing routes
app.post('/api/excel-to-word', upload.single('file'), asyncHandler(async (req, res) => {
  if (!req.file) {
    throw new AppError('No file uploaded', 400);
  }

  console.log('Processing Excel file:', req.file.filename);

  try {
    // Validate file size
    if (!ValidationUtils.validateFileSize(req.file.size, parseInt(process.env.MAX_FILE_SIZE_MB) || 10)) {
      throw new AppError('File size exceeds limit', 400);
    }

    // Process Excel file
    const excelData = await excelService.readExcelFile(req.file.path);
    
    // Generate Word document with tables
    const wordFilePath = await wordService.generateWordDocument(excelData, req.file.filename);
    
    // Send file
    res.download(wordFilePath, 'thesis-tables.docx', async (err) => {
      if (err) {
        console.error('Error sending file:', err);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error downloading file' });
        }
      }
      
      // Cleanup
      await cleanup.cleanupFiles([req.file.path, wordFilePath]);
    });

  } catch (error) {
    // Cleanup on error
    await cleanup.cleanupFiles([req.file.path]);
    throw error;
  }
}));

app.post('/api/word-to-pdf', upload.single('file'), asyncHandler(async (req, res) => {
  if (!req.file) {
    throw new AppError('No file uploaded', 400);
  }

  console.log('Processing Word file for ZIP generation:', req.file.filename);

  const graphFiles = [];

  try {
    // Validate file size
    if (!ValidationUtils.validateFileSize(req.file.size, parseInt(process.env.MAX_FILE_SIZE_MB) || 10)) {
      throw new AppError('File size exceeds limit', 400);
    }

    // Parse Word document
    const tables = await wordParserService.parseWordDocument(req.file.path);
    
    // Validate tables
    const validation = ValidationUtils.validateWordTables(tables);
    if (!validation.valid) {
      throw new AppError(`Invalid table format: ${validation.errors.join(', ')}`, 422);
    }
    
    // Generate graphs (individual PNG files)
    const graphs = await graphService.generateGraphs(tables);
    graphs.forEach(g => graphFiles.push(g.imagePath));
    
    // Create ZIP archive with individual PNG graphs
    const zipFilePath = await zipService.createZipArchive(graphs, req.file.filename);
    
    console.log('ZIP file created:', zipFilePath);
    
    // Send ZIP file
    res.download(zipFilePath, 'thesis-graphs.zip', async (err) => {
      if (err) {
        console.error('Error sending ZIP file:', err);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error downloading ZIP file' });
        }
      }
      
      // Cleanup all generated files
      const filesToClean = [req.file.path, zipFilePath, graphFiles];
      await cleanup.cleanupFiles(filesToClean);
    });

  } catch (error) {
    // Cleanup on error
    const filesToClean = [req.file.path, graphFiles];
    await cleanup.cleanupFiles(filesToClean);
    throw error;
  }
}));

// Signs & Symptoms Analysis
app.post('/api/signs-analysis', upload.single('file'), asyncHandler(async (req, res) => {
  if (!req.file) {
    throw new AppError('No file uploaded', 400);
  }

  console.log('Processing Signs & Symptoms file:', req.file.filename);

  const graphFiles = [];

  try {
    // Validate file size
    if (!ValidationUtils.validateFileSize(req.file.size, parseInt(process.env.MAX_FILE_SIZE_MB) || 10)) {
      throw new AppError('File size exceeds limit', 400);
    }

    // Parse Signs & Symptoms document
    const tables = await signsParserService.parseSignsDocument(req.file.path);
    
    // Validate signs data using ValidationUtils
    const validation = ValidationUtils.validateSignsDocument(tables);
    if (!validation.valid) {
      throw new AppError(`Invalid Signs document format: ${validation.errors.join(', ')}`, 422);
    }
    
    // Generate signs analysis graphs (individual PNG files)
    const graphs = await graphService.generateSignsGraphs(tables);
    graphs.forEach(g => graphFiles.push(g.imagePath));
    
    // Create ZIP archive with individual PNG graphs
    const zipFilePath = await zipService.createZipArchive(graphs, req.file.filename, 'signs');
    
    console.log('Signs analysis ZIP file created:', zipFilePath);
    
    // Send ZIP file
    res.download(zipFilePath, 'signs-symptoms-graphs.zip', async (err) => {
      if (err) {
        console.error('Error sending Signs ZIP file:', err);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error downloading Signs ZIP file' });
        }
      }
      
      // Cleanup all generated files
      const filesToClean = [req.file.path, zipFilePath, graphFiles];
      await cleanup.cleanupFiles(filesToClean);
    });

  } catch (error) {
    // Cleanup on error
    const filesToClean = [req.file.path, graphFiles];
    await cleanup.cleanupFiles(filesToClean);
    
    // Check for specific error types and provide appropriate messages
    if (error.message.includes('Failed to parse')) {
      throw new AppError(`Failed to generate Signs analysis: ${error.message}`, 500);
    }
    throw error;
  }
}));

// Health check endpoint
app.get('/api/health', asyncHandler(async (req, res) => {
  const uploadDirSize = await cleanup.getDirectorySize();
  
  res.json({ 
    status: 'OK', 
    message: 'Thesis Helper API is running',
    version: '1.0.0',
    environment: process.env.NODE_ENV || 'development',
    uploadDirectory: {
      size: uploadDirSize
    }
  });
}));

// Admin endpoints (protect these in production)
app.get('/api/admin/cleanup', asyncHandler(async (req, res) => {
  const deletedCount = await cleanup.cleanOldFiles();
  res.json({ 
    success: true, 
    message: `Cleanup completed. Deleted ${deletedCount} files.` 
  });
}));

app.get('/api/admin/stats', asyncHandler(async (req, res) => {
  const dirSize = await cleanup.getDirectorySize();
  res.json({ 
    success: true, 
    uploadDirectory: dirSize
  });
}));

// 404 handler
app.use(notFound);

// Error handling middleware
app.use(errorHandler);

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM signal received: closing HTTP server');
  cleanup.stopPeriodicCleanup();
  server.close(() => {
    console.log('HTTP server closed');
  });
});

process.on('SIGINT', () => {
  console.log('SIGINT signal received: closing HTTP server');
  cleanup.stopPeriodicCleanup();
  server.close(() => {
    console.log('HTTP server closed');
  });
});

// Start server
const server = app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
  console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
  console.log(`Upload directory: ${process.env.UPLOAD_DIR || 'uploads'}`);
  console.log(`Max file size: ${process.env.MAX_FILE_SIZE_MB || 10}MB`);
});

module.exports = app;