const fs = require('fs').promises;
const path = require('path');

class CleanupUtility {
  constructor(uploadDir = 'uploads', retentionHours = 24) {
    this.uploadDir = uploadDir;
    this.retentionHours = retentionHours;
  }

  // Clean up old files
  async cleanOldFiles() {
    try {
      const files = await fs.readdir(this.uploadDir);
      const now = Date.now();
      const maxAge = this.retentionHours * 60 * 60 * 1000; // Convert hours to milliseconds
      
      let deletedCount = 0;
      
      for (const file of files) {
        const filePath = path.join(this.uploadDir, file);
        
        try {
          const stats = await fs.stat(filePath);
          
          // Check if file is older than retention period
          if (now - stats.mtimeMs > maxAge) {
            await fs.unlink(filePath);
            deletedCount++;
            console.log(`Deleted old file: ${file}`);
          }
        } catch (error) {
          console.error(`Error processing file ${file}:`, error);
        }
      }
      
      console.log(`Cleanup completed. Deleted ${deletedCount} files.`);
      return deletedCount;
      
    } catch (error) {
      console.error('Error during cleanup:', error);
      throw error;
    }
  }

  // Clean specific file types
  async cleanFilesByType(extensions) {
    try {
      const files = await fs.readdir(this.uploadDir);
      let deletedCount = 0;
      
      for (const file of files) {
        const ext = path.extname(file).toLowerCase();
        
        if (extensions.includes(ext)) {
          const filePath = path.join(this.uploadDir, file);
          
          try {
            await fs.unlink(filePath);
            deletedCount++;
            console.log(`Deleted file: ${file}`);
          } catch (error) {
            console.error(`Error deleting file ${file}:`, error);
          }
        }
      }
      
      return deletedCount;
      
    } catch (error) {
      console.error('Error during cleanup by type:', error);
      throw error;
    }
  }

  // Clean all temporary files
  async cleanAllTempFiles() {
    const tempExtensions = ['.tmp', '.temp', '~'];
    return await this.cleanFilesByType(tempExtensions);
  }

  // Get directory size
  async getDirectorySize() {
    try {
      const files = await fs.readdir(this.uploadDir);
      let totalSize = 0;
      
      for (const file of files) {
        const filePath = path.join(this.uploadDir, file);
        const stats = await fs.stat(filePath);
        totalSize += stats.size;
      }
      
      return {
        bytes: totalSize,
        mb: (totalSize / (1024 * 1024)).toFixed(2)
      };
      
    } catch (error) {
      console.error('Error calculating directory size:', error);
      throw error;
    }
  }

  // Schedule periodic cleanup
  startPeriodicCleanup(intervalHours = 6) {
    // Run cleanup immediately
    this.cleanOldFiles().catch(console.error);
    
    // Schedule periodic cleanup
    const intervalMs = intervalHours * 60 * 60 * 1000;
    
    this.cleanupInterval = setInterval(() => {
      this.cleanOldFiles().catch(console.error);
    }, intervalMs);
    
    console.log(`Periodic cleanup scheduled every ${intervalHours} hours`);
  }

  // Stop periodic cleanup
  stopPeriodicCleanup() {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
      this.cleanupInterval = null;
      console.log('Periodic cleanup stopped');
    }
  }

  // Clean up specific files
  async cleanupFiles(filePaths) {
    const results = [];
    
    for (const filePath of filePaths) {
      try {
        await fs.unlink(filePath);
        results.push({ file: filePath, success: true });
      } catch (error) {
        results.push({ file: filePath, success: false, error: error.message });
      }
    }
    
    return results;
  }

  // Ensure upload directory exists
  async ensureUploadDirectory() {
    try {
      await fs.mkdir(this.uploadDir, { recursive: true });
      console.log(`Upload directory ensured: ${this.uploadDir}`);
    } catch (error) {
      console.error('Error creating upload directory:', error);
      throw error;
    }
  }
}

module.exports = CleanupUtility;