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
      'word-to-pdf': ['.docx', '.doc']
    };
    
    const route = req.path.includes('excel-to-word') ? 'excel-to-word' : 'word-to-pdf';
    const ext = path.extname(file.originalname).toLowerCase();
    
    if (!ValidationUtils.validateFileExtension(file.originalname, allowedTypes[route])) {
      cb(new AppError(`Invalid file type. Allowed types: ${allowedTypes[route].join(', ')}`, 400));
    } else {
      cb(null, true);
    }
  }
});

// Routes
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

// UPDATED: This endpoint now generates ZIP with individual PNG graphs instead of PDF
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
    
    // Create ZIP archive with individual PNG graphs instead of PDF
    const zipFilePath = await zipService.createZipArchive(graphs, req.file.filename);
    
    console.log('ZIP file created:', zipFilePath);
    
    // Send ZIP file instead of PDF
    res.download(zipFilePath, 'thesis-graphs.zip', async (err) => {
      if (err) {
        console.error('Error sending ZIP file:', err);
        if (!res.headersSent) {
          res.status(500).json({ success: false, error: 'Error downloading ZIP file' });
        }
      }
      
      // Cleanup all generated files
      const filesToClean = [req.file.path, zipFilePath, ...graphFiles];
      await cleanup.cleanupFiles(filesToClean);
    });

  } catch (error) {
    // Cleanup on error
    const filesToClean = [req.file.path, ...graphFiles];
    await cleanup.cleanupFiles(filesToClean);
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