# Thesis Helper Testing Guide

## Prerequisites
- Node.js installed
- Both backend and frontend servers running
- Sample Excel and Word files for testing

## Quick Test Setup

### 1. Generate Test Data
```bash
cd backend
node createSampleExcel.js
```
This creates `sample-masterchart.xlsx` for testing.

### 2. Start Servers
Terminal 1 (Backend):
```bash
cd backend
npm run dev
```

Terminal 2 (Frontend):
```bash
cd frontend
npm start
```

## Manual Testing Checklist

### 1. Excel to Word Conversion

#### Test Case 1.1: Valid Excel File
1. Open http://localhost:3000
2. Click "Excel to Word Tables"
3. Upload `sample-masterchart.xlsx`
4. Click "Generate Word Document"
5. **Expected:** Download starts for `thesis-tables.docx`
6. **Verify:** 
   - Word file contains all tables
   - Each table has correct headers
   - Total row is present and bold
   - Percentages sum to 100%

#### Test Case 1.2: Invalid File Type
1. Try uploading a .txt file
2. **Expected:** Error message "Invalid file type"

#### Test Case 1.3: Missing Sheets
1. Create Excel with only one sheet
2. Upload file
3. **Expected:** Error "Missing Trial Group or Control Group sheets"

#### Test Case 1.4: Large File
1. Upload file > 10MB
2. **Expected:** Error "File size too large"

### 2. Word to PDF Conversion

#### Test Case 2.1: Valid Word File
1. Use the generated `thesis-tables.docx`
2. Click "Word Tables to PDF Graphs"
3. Upload the Word file
4. Click "Generate PDF"
5. **Expected:** Download starts for `thesis-graphs.pdf`
6. **Verify:**
   - PDF contains all graphs
   - Graphs have 3D effect
   - Labels are readable
   - Title is present

#### Test Case 2.2: Empty Word File
1. Upload empty Word document
2. **Expected:** Error "No tables found"

#### Test Case 2.3: Invalid Table Format
1. Create Word with incorrectly formatted tables
2. **Expected:** Error describing the issue

### 3. API Testing

#### Test Case 3.1: Health Check
```bash
curl http://localhost:5000/api/health
```
**Expected:** JSON with status "OK"

#### Test Case 3.2: Direct API Excel Upload
```bash
curl -X POST \
  http://localhost:5000/api/excel-to-word \
  -F 'file=@sample-masterchart.xlsx' \
  --output thesis-tables.docx
```
**Expected:** Word file downloads successfully

#### Test Case 3.3: Missing File
```bash
curl -X POST http://localhost:5000/api/excel-to-word
```
**Expected:** Error "No file uploaded"

### 4. Error Handling Tests

#### Test Case 4.1: Network Error
1. Stop backend server
2. Try uploading a file
3. **Expected:** Error message about connection

#### Test Case 4.2: Concurrent Uploads
1. Upload multiple files simultaneously
2. **Expected:** All uploads process successfully

### 5. Edge Cases

#### Test Case 5.1: Special Characters
1. Create Excel with special characters in data
2. Upload and process
3. **Expected:** Characters handled correctly

#### Test Case 5.2: Empty Columns
1. Create Excel with some empty columns
2. **Expected:** Empty columns skipped

#### Test Case 5.3: Large Dataset
1. Create Excel with 1000+ rows
2. **Expected:** Processes successfully (may take time)

## Automated Testing Setup

### Backend Unit Tests
Create `backend/tests/services.test.js`:

```javascript
const excelService = require('../services/excelService');
const ValidationUtils = require('../utils/validation');

describe('Excel Service Tests', () => {
  test('should read valid Excel file', async () => {
    const data = await excelService.readExcelFile('sample-masterchart.xlsx');
    expect(data.tables).toBeDefined();
    expect(data.tables.length).toBeGreaterThan(0);
  });
  
  test('should validate file extensions', () => {
    expect(ValidationUtils.validateFileExtension('test.xlsx', ['.xlsx'])).toBe(true);
    expect(ValidationUtils.validateFileExtension('test.pdf', ['.xlsx'])).toBe(false);
  });
});
```

Run tests:
```bash
cd backend
npm test
```

### Frontend Component Tests
Create `frontend/src/App.test.js`:

```javascript
import { render, screen, fireEvent } from '@testing-library/react';
import App from './App';

test('renders main title', () => {
  render(<App />);
  const title = screen.getByText(/Thesis Helper/i);
  expect(title).toBeInTheDocument();
});

test('shows two options initially', () => {
  render(<App />);
  expect(screen.getByText(/Excel to Word Tables/i)).toBeInTheDocument();
  expect(screen.getByText(/Word Tables to PDF Graphs/i)).toBeInTheDocument();
});

test('clicking option shows upload screen', () => {
  render(<App />);
  fireEvent.click(screen.getByText(/Excel to Word Tables/i));
  expect(screen.getByText(/Back to options/i)).toBeInTheDocument();
});
```

## Performance Testing

### Load Test Script
Create `backend/tests/load-test.js`:

```javascript
const fs = require('fs');
const FormData = require('form-data');
const axios = require('axios');

async function loadTest(concurrent = 5) {
  const promises = [];
  
  for (let i = 0; i < concurrent; i++) {
    const form = new FormData();
    form.append('file', fs.createReadStream('sample-masterchart.xlsx'));
    
    promises.push(
      axios.post('http://localhost:5000/api/excel-to-word', form, {
        headers: form.getHeaders()
      })
    );
  }
  
  console.time('Load Test');
  await Promise.all(promises);
  console.timeEnd('Load Test');
}

loadTest(10).catch(console.error);
```

## Browser Compatibility Testing

Test on:
- Chrome (latest)
- Firefox (latest)
- Safari (latest)
- Edge (latest)
- Mobile browsers

## Security Testing

### Test Case S1: Path Traversal
1. Try uploading file with name `../../etc/passwd`
2. **Expected:** Filename sanitized

### Test Case S2: Large Payload
1. Try uploading file > 20MB
2. **Expected:** Rejected by server

### Test Case S3: Malformed Data
1. Send malformed multipart data
2. **Expected:** Proper error response

## Debugging Tips

### Enable Debug Logging
Add to backend `.env`:
```
DEBUG=*
LOG_LEVEL=debug
```

### Check File Processing
```javascript
// Add to services for debugging
console.log('Processing file:', filename);
console.log('Sheet names:', workbook.worksheets.map(ws => ws.name));
console.log('Row count:', worksheet.rowCount);
```

### Monitor Memory Usage
```javascript
// Add to server.js
setInterval(() => {
  const used = process.memoryUsage();
  console.log('Memory Usage:', {
    rss: `${Math.round(used.rss / 1024 / 1024)}MB`,
    heapTotal: `${Math.round(used.heapTotal / 1024 / 1024)}MB`,
    heapUsed: `${Math.round(used.heapUsed / 1024 / 1024)}MB`
  });
}, 30000);
```

## Common Issues and Solutions

### Issue 1: Canvas Installation Fails
**Solution:**
```bash
# Ubuntu/Debian
sudo apt-get install build-essential libcairo2-dev libpango1.0-dev libjpeg-dev libgif-dev librsvg2-dev

# macOS
brew install pkg-config cairo pango libpng jpeg giflib librsvg

# Then reinstall
npm install canvas
```

### Issue 2: File Upload Timeout
**Solution:** Increase timeout in nginx.conf or server configuration

### Issue 3: Memory Issues with Large Files
**Solution:** Implement streaming for large files or increase Node.js memory:
```bash
node --max-old-space-size=4096 server.js
```

## Production Testing Checklist

- [ ] All endpoints return correct status codes
- [ ] File cleanup works after specified retention period
- [ ] Error messages don't expose sensitive information
- [ ] CORS is properly configured
- [ ] File size limits are enforced
- [ ] Concurrent requests handled properly
- [ ] Memory usage stays within limits
- [ ] Response times are acceptable
- [ ] Downloads work on all browsers
- [ ] Mobile responsiveness works correctly