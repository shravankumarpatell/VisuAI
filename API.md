# Thesis Helper API Documentation

## Base URL
- Development: `http://localhost:5000/api`
- Production: `https://your-domain.com/api`

## Authentication
Currently, the API does not require authentication. In production, consider implementing API keys or JWT tokens.

## Endpoints

### 1. Excel to Word Conversion
Convert Excel master charts to Word document with formatted tables.

**Endpoint:** `POST /api/excel-to-word`

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body:
  - `file`: Excel file (.xlsx or .xls)

**Example Request:**
```bash
curl -X POST \
  http://localhost:5000/api/excel-to-word \
  -H 'Content-Type: multipart/form-data' \
  -F 'file=@master-chart.xlsx'
```

**Response:**
- Success: Word document file download (thesis-tables.docx)
- Error: JSON object with error details

**Error Responses:**
```json
{
  "success": false,
  "error": "No file uploaded"
}
```

```json
{
  "success": false,
  "error": "Invalid file type. Allowed types: .xlsx, .xls"
}
```

```json
{
  "success": false,
  "error": "Missing \"Trial Group\" sheet"
}
```

### 2. Word to PDF Conversion
Convert Word document tables to PDF with 3D graphs.

**Endpoint:** `POST /api/word-to-pdf`

**Request:**
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body:
  - `file`: Word file (.docx or .doc)

**Example Request:**
```bash
curl -X POST \
  http://localhost:5000/api/word-to-pdf \
  -H 'Content-Type: multipart/form-data' \
  -F 'file=@tables.docx'
```

**Response:**
- Success: PDF file download (thesis-graphs.pdf)
- Error: JSON object with error details

**Error Responses:**
```json
{
  "success": false,
  "error": "No tables found in the Word document"
}
```

```json
{
  "success": false,
  "error": "Invalid table format: Table 1 is missing required field: trialGroup"
}
```

### 3. Health Check
Check API status and server information.

**Endpoint:** `GET /api/health`

**Request:**
- Method: `GET`

**Example Request:**
```bash
curl http://localhost:5000/api/health
```

**Response:**
```json
{
  "status": "OK",
  "message": "Thesis Helper API is running",
  "version": "1.0.0",
  "environment": "development",
  "uploadDirectory": {
    "size": {
      "bytes": 1048576,
      "mb": "1.00"
    }
  }
}
```

### 4. Admin - Manual Cleanup
Trigger manual cleanup of old files (protect in production).

**Endpoint:** `GET /api/admin/cleanup`

**Request:**
- Method: `GET`

**Example Request:**
```bash
curl http://localhost:5000/api/admin/cleanup
```

**Response:**
```json
{
  "success": true,
  "message": "Cleanup completed. Deleted 5 files."
}
```

### 5. Admin - Statistics
Get upload directory statistics (protect in production).

**Endpoint:** `GET /api/admin/stats`

**Request:**
- Method: `GET`

**Example Request:**
```bash
curl http://localhost:5000/api/admin/stats
```

**Response:**
```json
{
  "success": true,
  "uploadDirectory": {
    "bytes": 5242880,
    "mb": "5.00"
  }
}
```

## File Format Requirements

### Excel File Format
1. **Required Sheets:**
   - "Trial Group" - Contains trial group data
   - "Control Group" - Contains control group data

2. **Sheet Structure:**
   - Row 1: Headers
   - Column 1: Sl.No (will be skipped)
   - Column 2: OPD No (will be skipped)
   - Column 3+: Data columns (Age, Gender, Religion, etc.)

3. **Example Structure:**
   | Sl.No | OPD No | Age   | Gender | Religion | Occupation |
   |-------|--------|-------|--------|----------|------------|
   | 1     | OPD001 | 20-30 | Male   | Hindu    | Student    |
   | 2     | OPD002 | 31-40 | Female | Muslim   | Teacher    |

### Word File Format
1. **Table Structure:**
   - Must contain properly formatted tables
   - Each table should have 5 columns

2. **Required Columns:**
   | Column Name     | Description              |
   |----------------|--------------------------|
   | Category       | Row category/label       |
   | Trial Group    | Trial group count        |
   | Control Group  | Control group count      |
   | Total          | Sum of trial + control   |
   | Percentage (%) | Percentage of total      |

3. **Example Table:**
   | Age   | Trial Group | Control Group | Total | Percentage (%) |
   |-------|-------------|---------------|-------|----------------|
   | 20-30 | 6           | 4             | 10    | 25.00%         |
   | 31-40 | 3           | 3             | 6     | 15.00%         |
   | Total | 20          | 20            | 40    | 100.00%        |

## Error Handling

### Common Error Codes
- `400` - Bad Request (invalid file, missing parameters)
- `422` - Unprocessable Entity (invalid file content/format)
- `500` - Internal Server Error

### Error Response Format
```json
{
  "success": false,
  "error": "Error message description",
  "stack": "Stack trace (only in development mode)"
}
```

## Rate Limiting
- Default: 100 requests per 15 minutes per IP
- File upload endpoints: 10 requests per 15 minutes per IP

## File Size Limits
- Maximum file size: 10MB (configurable via MAX_FILE_SIZE_MB)
- Supported formats:
  - Excel: .xlsx, .xls
  - Word: .docx, .doc

## Best Practices

1. **File Validation:**
   - Always validate file format before uploading
   - Check file size client-side to avoid unnecessary uploads

2. **Error Handling:**
   - Implement proper error handling on the client
   - Show user-friendly error messages

3. **Performance:**
   - Large files may take time to process
   - Implement progress indicators on the client

4. **Security:**
   - In production, implement authentication
   - Use HTTPS for all API calls
   - Validate and sanitize all inputs

## Examples

### JavaScript/Fetch Example
```javascript
// Excel to Word
const uploadExcel = async (file) => {
  const formData = new FormData();
  formData.append('file', file);
  
  try {
    const response = await fetch('http://localhost:5000/api/excel-to-word', {
      method: 'POST',
      body: formData
    });
    
    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error);
    }
    
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'thesis-tables.docx';
    a.click();
  } catch (error) {
    console.error('Upload failed:', error);
  }
};
```

### Python Example
```python
import requests

# Excel to Word
def convert_excel_to_word(file_path):
    url = 'http://localhost:5000/api/excel-to-word'
    
    with open(file_path, 'rb') as f:
        files = {'file': f}
        response = requests.post(url, files=files)
    
    if response.status_code == 200:
        with open('thesis-tables.docx', 'wb') as f:
            f.write(response.content)
        print('File saved successfully')
    else:
        print('Error:', response.json())
```

## Deployment Notes

1. **Environment Variables:**
   - Set all required environment variables
   - Use different values for production

2. **File Storage:**
   - Consider using cloud storage for production
   - Implement proper backup strategies

3. **Monitoring:**
   - Implement logging and monitoring
   - Track API usage and errors

4. **Security:**
   - Enable CORS restrictions
   - Implement rate limiting
   - Add authentication for admin endpoints