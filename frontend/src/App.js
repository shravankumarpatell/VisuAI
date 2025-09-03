import React, { useState } from 'react';

const API_URL = 'http://localhost:5000/api';

function App() {
  const [activeOption, setActiveOption] = useState(null);
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    setFile(selectedFile);
    setError('');
    setSuccess('');
  };

  const handleExcelToWord = async () => {
    if (!file) {
      setError('Please select an Excel file');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      const response = await fetch(`${API_URL}/excel-to-word`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate Word document');
      }

      const blob = await response.blob();
      
      // Create download link
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'thesis-tables.docx');
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      setSuccess('Word document generated successfully!');
      setFile(null);
      document.getElementById('file-input').value = '';
    } catch (err) {
      console.error('Error:', err);
      setError(err.message || 'Failed to generate Word document');
    } finally {
      setLoading(false);
    }
  };

  // NEW: Master Chart Analysis function
  const handleMasterChartAnalysis = async () => {
    if (!file) {
      setError('Please select a Master Chart Excel file');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      const response = await fetch(`${API_URL}/master-chart-analysis`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate Master Chart analysis');
      }

      const blob = await response.blob();
      
      // Create download link
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'master-chart-analysis.docx');
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      setSuccess('Master Chart statistical analysis generated successfully!');
      setFile(null);
      document.getElementById('file-input').value = '';
    } catch (err) {
      console.error('Error:', err);
      setError(err.message || 'Failed to generate Master Chart analysis');
    } finally {
      setLoading(false);
    }
  };

  const handleWordToPdf = async () => {
    if (!file) {
      setError('Please select a Word file');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      const response = await fetch(`${API_URL}/word-to-pdf`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate graphs ZIP');
      }

      const blob = await response.blob();
      
      // Create download link for ZIP file
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'thesis-graphs.zip');
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      setSuccess('ZIP file with individual PNG graphs generated successfully!');
      setFile(null);
      document.getElementById('file-input').value = '';
    } catch (err) {
      console.error('Error:', err);
      setError(err.message || 'Failed to generate graphs ZIP');
    } finally {
      setLoading(false);
    }
  };

  const handleSignsAnalysis = async () => {
    if (!file) {
      setError('Please select a Word file with Signs & Symptoms data');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      const response = await fetch(`${API_URL}/signs-analysis`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate Signs & Symptoms graphs');
      }

      const blob = await response.blob();
      
      // Create download link for ZIP file
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'signs-symptoms-graphs.zip');
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      setSuccess('Signs & Symptoms analysis graphs generated successfully!');
      setFile(null);
      document.getElementById('file-input').value = '';
    } catch (err) {
      console.error('Error:', err);
      setError(err.message || 'Failed to generate Signs & Symptoms graphs');
    } finally {
      setLoading(false);
    }
  };

  // NEW: Improvement Analysis function
  const handleImprovementAnalysis = async () => {
    if (!file) {
      setError('Please select a Word file with improvement percentage table');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      const response = await fetch(`${API_URL}/improvement-analysis`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate improvement analysis graph');
      }

      const blob = await response.blob();
      
      // Create download link for PNG file
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'improvement-analysis-graph.png');
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);

      setSuccess('Improvement analysis graph generated successfully!');
      setFile(null);
      document.getElementById('file-input').value = '';
    } catch (err) {
      console.error('Error:', err);
      setError(err.message || 'Failed to generate improvement analysis graph');
    } finally {
      setLoading(false);
    }
  };

  const resetSelection = () => {
    setActiveOption(null);
    setFile(null);
    setError('');
    setSuccess('');
    const fileInput = document.getElementById('file-input');
    if (fileInput) fileInput.value = '';
  };

  const getOptionConfig = (option) => {
    const configs = {
      'excel-to-word': {
        title: 'Excel to Word Tables',
        description: 'Upload Excel master charts to generate formatted tables in Word',
        icon: '📊',
        fileTypes: '.xlsx,.xls',
        fileDescription: 'Accepted formats: .xlsx, .xls',
        buttonText: 'Generate Word Document',
        handler: handleExcelToWord
      },
      'master-chart-analysis': {
        title: 'Master Chart Statistical Analysis',
        description: 'Upload Excel master chart with BT/AT data to generate comprehensive statistical analysis tables',
        icon: '📊',
        fileTypes: '.xlsx,.xls',
        fileDescription: 'Accepted formats: .xlsx, .xls (Must contain BT/AT columns)',
        buttonText: 'Generate Statistical Analysis',
        handler: handleMasterChartAnalysis
      },
      'word-to-pdf': {
        title: 'Word Tables to PNG Graphs ZIP',
        description: 'Upload Word document with tables to generate individual 3D PNG graphs in ZIP file',
        icon: '📈',
        fileTypes: '.docx,.doc',
        fileDescription: 'Accepted formats: .docx, .doc',
        buttonText: 'Generate PNG Graphs ZIP',
        handler: handleWordToPdf
      },
      'signs-analysis': {
        title: 'Signs & Symptoms Analysis',
        description: 'Upload Word document with Signs & Symptoms tables to generate percentage-based assessment graphs',
        icon: '🔬',
        fileTypes: '.docx,.doc',
        fileDescription: 'Accepted formats: .docx, .doc',
        buttonText: 'Generate Signs Analysis Graphs',
        handler: handleSignsAnalysis
      },
      'improvement-analysis': {
        title: 'Improvement Analysis Graph',
        description: 'Upload Word document with improvement percentage table to generate 3D visualization graph',
        icon: '📈',
        fileTypes: '.docx,.doc',
        fileDescription: 'Accepted formats: .docx, .doc (Must contain improvement percentage table)',
        buttonText: 'Generate Improvement Graph',
        handler: handleImprovementAnalysis
      }
    };
    return configs[option];
  };

  return (
    <div className="app">
      <header className="app-header">
        <h1>Thesis Helper</h1>
        <p>Generate tables and graphs for your medical thesis</p>
      </header>

      <main className="app-main">
        {!activeOption ? (
          <div className="options-container">
            <h2>Choose an option:</h2>
            <div className="option-cards">
              <div 
                className="option-card"
                onClick={() => setActiveOption('excel-to-word')}
              >
                <div className="option-icon">📊</div>
                <h3>Excel to Word Tables</h3>
                <p>Upload Excel master charts to generate formatted tables in Word</p>
              </div>
              
              {/* Master Chart Analysis Option */}
              <div 
                className="option-card"
                onClick={() => setActiveOption('master-chart-analysis')}
              >
                <div className="option-icon">📊</div>
                <h3>Master Chart Statistical Analysis</h3>
                <p>Upload Excel master chart with BT/AT data to generate comprehensive statistical analysis with t-tests, p-values, and significance levels</p>
              </div>
              
              <div 
                className="option-card"
                onClick={() => setActiveOption('word-to-pdf')}
              >
                <div className="option-icon">📈</div>
                <h3>Word Tables to PNG Graphs ZIP</h3>
                <p>Upload Word document with tables to generate individual 3D PNG graphs in ZIP file</p>
              </div>

              <div 
                className="option-card"
                onClick={() => setActiveOption('signs-analysis')}
              >
                <div className="option-icon">🔬</div>
                <h3>Signs & Symptoms Analysis</h3>
                <p>Upload Word document with Signs & Symptoms tables to generate percentage-based assessment graphs</p>
              </div>

              {/* NEW: Improvement Analysis Option */}
              <div 
                className="option-card featured"
                onClick={() => setActiveOption('improvement-analysis')}
              >
                <div className="option-icon">📈</div>
                <div className="option-badge">NEW</div>
                <h3>Improvement Analysis Graph</h3>
                <p>Upload Word document with improvement percentage table to generate 3D bar chart visualization</p>
              </div>
            </div>
          </div>
        ) : (
          <div className="upload-container">
            <button className="back-button" onClick={resetSelection}>
              ← Back to options
            </button>
            
            <h2>{getOptionConfig(activeOption).title}</h2>
            
            {activeOption === 'master-chart-analysis' && (
              <div className="info-panel">
                <h4>📋 Master Chart Requirements:</h4>
                <ul>
                  <li>Excel file with multiple sheets containing patient data</li>
                  <li>Each sheet should have columns for Before Treatment (BT) and After Treatment (AT) values</li>
                  <li>Column headers like "Vedana BT", "Vedana AT", "Varna BT", "Varna AT", etc.</li>
                  <li>Numerical data in BT/AT columns for statistical analysis</li>
                  <li>System will automatically detect parameter pairs and calculate:</li>
                  <ul>
                    <li>Paired t-test statistics</li>
                    <li>P-values and significance levels</li>
                    <li>Effectiveness percentages</li>
                    <li>Mean ± Standard deviation</li>
                  </ul>
                </ul>
              </div>
            )}
            
            {activeOption === 'improvement-analysis' && (
              <div className="info-panel">
                <h4>📋 Improvement Analysis Requirements:</h4>
                <ul>
                  <li>Word document with a single improvement percentage table</li>
                  <li>Table should have improvement categories (rows): Cured(100%), Marked improved(75-100%), etc.</li>
                  <li>Columns should show Group A and Group B data for different time periods (7th day, 14th day, 21st day, 28th day)</li>
                  <li>Data format: count (percentage%) - e.g., "8 (40%)"</li>
                  <li>System will generate a 3D bar chart visualization showing:</li>
                  <ul>
                    <li>Improvement categories on X-axis</li>
                    <li>Patient count on Y-axis</li>
                    <li>Different colored bars for each time period</li>
                    <li>Separate grouping for Group A and Group B</li>
                  </ul>
                </ul>
              </div>
            )}
            
            <div className="upload-area">
              <input
                id="file-input"
                type="file"
                accept={getOptionConfig(activeOption).fileTypes}
                onChange={handleFileChange}
                className="file-input"
              />
              
              <label htmlFor="file-input" className="file-label">
                <div className="upload-icon">📁</div>
                <p>
                  {file 
                    ? `Selected: ${file.name}` 
                    : `Click to select ${
                        activeOption === 'excel-to-word' || activeOption === 'master-chart-analysis' 
                          ? 'Excel' : 'Word'
                      } file`}
                </p>
                <small>{getOptionConfig(activeOption).fileDescription}</small>
              </label>
            </div>

            {error && (
              <div className="message error-message">
                <div className="error-icon">⚠️</div>
                <div className="error-content">
                  <p className="error-text">{error}</p>
                  {error.includes('Invalid master chart format') && (
                    <div className="error-help">
                      <p><strong>How to fix Master Chart format:</strong></p>
                      <ol>
                        <li>Ensure your Excel file has at least one sheet with data</li>
                        <li>Check that columns are named with BT/AT patterns (e.g., "Vedana BT", "Vedana AT")</li>
                        <li>Verify that BT/AT columns contain numerical values</li>
                        <li>Make sure you have at least 3 data rows for statistical validity</li>
                      </ol>
                    </div>
                  )}
                  {error.includes('Trial Group or Control Group') && (
                    <div className="error-help">
                      <p><strong>How to fix:</strong></p>
                      <ol>
                        <li>Open your Excel file</li>
                        <li>Rename first sheet to "Trial Group"</li>
                        <li>Rename second sheet to "Control Group"</li>
                        <li>Save and try again</li>
                      </ol>
                      <p className="error-tip">💡 Tip: The system now also accepts any Excel with at least 2 sheets, using them as Trial and Control groups automatically.</p>
                    </div>
                  )}
                  {error.includes('improvement percentage table') && (
                    <div className="error-help">
                      <p><strong>How to fix Improvement Analysis format:</strong></p>
                      <ol>
                        <li>Ensure your Word file contains a single table with improvement data</li>
                        <li>Check that the table has improvement categories as row headers</li>
                        <li>Verify columns for Group A and Group B with time periods (7th, 14th, 21st, 28th day)</li>
                        <li>Data should be in format: "count (percentage%)" - e.g., "8 (40%)"</li>
                      </ol>
                    </div>
                  )}
                </div>
              </div>
            )}

            {success && (
              <div className="message success-message">
                ✅ {success}
              </div>
            )}

            <button
              className="process-button"
              onClick={getOptionConfig(activeOption).handler}
              disabled={!file || loading}
            >
              {loading ? (
                <>
                  <span className="spinner"></span>
                  Processing...
                </>
              ) : (
                getOptionConfig(activeOption).buttonText
              )}
            </button>
          </div>
        )}
      </main>

      <footer className="app-footer">
        <p>Thesis Helper © 2024 | For MD Students</p>
      </footer>

      <style jsx>{`
        .app {
          min-height: 100vh;
          display: flex;
          flex-direction: column;
          background-color: #f5f7fa;
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        }

        .app-header {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          padding: 2rem;
          text-align: center;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .app-header h1 {
          margin: 0;
          font-size: 2.5rem;
          font-weight: 700;
        }

        .app-header p {
          margin: 0.5rem 0 0;
          font-size: 1.1rem;
          opacity: 0.9;
        }

        .app-main {
          flex: 1;
          padding: 2rem;
          max-width: 1200px;
          margin: 0 auto;
          width: 100%;
        }

        .options-container h2 {
          text-align: center;
          color: #333;
          margin-bottom: 2rem;
          font-size: 1.8rem;
        }

        .option-cards {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
          gap: 2rem;
          margin-top: 2rem;
        }

        .option-card {
          background: white;
          border-radius: 12px;
          padding: 2rem;
          text-align: center;
          cursor: pointer;
          transition: all 0.3s ease;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          border: 2px solid transparent;
          position: relative;
        }

        .option-card:hover {
          transform: translateY(-5px);
          box-shadow: 0 8px 20px rgba(0,0,0,0.15);
          border-color: #667eea;
        }

        .option-card.featured {
          border-color: #4CAF50;
          box-shadow: 0 4px 12px rgba(76, 175, 80, 0.2);
        }

        .option-card.featured:hover {
          border-color: #45a049;
          box-shadow: 0 8px 20px rgba(76, 175, 80, 0.3);
        }

        .option-badge {
          position: absolute;
          top: -10px;
          right: -10px;
          background: #4CAF50;
          color: white;
          padding: 0.25rem 0.75rem;
          border-radius: 12px;
          font-size: 0.75rem;
          font-weight: bold;
          transform: rotate(15deg);
        }

        .option-icon {
          font-size: 4rem;
          margin-bottom: 1rem;
        }

        .option-card h3 {
          color: #333;
          margin: 1rem 0;
          font-size: 1.5rem;
        }

        .option-card p {
          color: #666;
          margin: 0;
          line-height: 1.6;
        }

        .upload-container {
          background: white;
          border-radius: 12px;
          padding: 2rem;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          max-width: 700px;
          margin: 0 auto;
        }

        .info-panel {
          background: #f0f8ff;
          border: 1px solid #b3d9ff;
          border-radius: 8px;
          padding: 1.5rem;
          margin: 1.5rem 0;
        }

        .info-panel h4 {
          margin: 0 0 1rem;
          color: #0066cc;
          font-size: 1.1rem;
        }

        .info-panel ul {
          margin: 0;
          padding-left: 1.5rem;
          color: #333;
        }

        .info-panel li {
          margin: 0.5rem 0;
          line-height: 1.4;
        }

        .info-panel ul ul {
          margin-top: 0.25rem;
        }

        .back-button {
          background: none;
          border: none;
          color: #667eea;
          cursor: pointer;
          font-size: 1rem;
          padding: 0;
          margin-bottom: 1rem;
          display: flex;
          align-items: center;
          transition: color 0.3s ease;
          font-family: inherit;
        }

        .back-button:hover {
          color: #764ba2;
        }

        .upload-container h2 {
          text-align: center;
          color: #333;
          margin-bottom: 2rem;
        }

        .upload-area {
          margin: 2rem 0;
        }

        .file-input {
          display: none;
        }

        .file-label {
          display: block;
          padding: 3rem;
          border: 2px dashed #ddd;
          border-radius: 8px;
          text-align: center;
          cursor: pointer;
          transition: all 0.3s ease;
          background: #f8f9fa;
        }

        .file-label:hover {
          border-color: #667eea;
          background: #f0f2ff;
        }

        .upload-icon {
          font-size: 3rem;
          margin-bottom: 1rem;
        }

        .file-label p {
          margin: 0.5rem 0;
          color: #333;
          font-size: 1.1rem;
        }

        .file-label small {
          color: #666;
          font-size: 0.9rem;
        }

        .message {
          padding: 1rem;
          border-radius: 6px;
          margin: 1rem 0;
          text-align: center;
        }

        .error-message {
          background: #fee;
          color: #c33;
          border: 1px solid #fcc;
          display: flex;
          align-items: flex-start;
          gap: 1rem;
          text-align: left;
        }

        .error-icon {
          font-size: 1.5rem;
          flex-shrink: 0;
        }

        .error-content {
          flex: 1;
        }

        .error-text {
          margin: 0 0 0.5rem;
          font-weight: 500;
        }

        .error-help {
          margin-top: 0.5rem;
          font-size: 0.9rem;
          color: #666;
        }

        .error-help ol {
          margin: 0.5rem 0;
          padding-left: 1.5rem;
        }

        .error-help li {
          margin: 0.25rem 0;
        }

        .error-tip {
          margin-top: 0.5rem;
          padding: 0.5rem;
          background: #fff;
          border-radius: 4px;
          font-size: 0.85rem;
        }

        .success-message {
          background: #efe;
          color: #3c3;
          border: 1px solid #cfc;
        }

        .process-button {
          width: 100%;
          padding: 1rem;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          border: none;
          border-radius: 6px;
          font-size: 1.1rem;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.3s ease;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 0.5rem;
          font-family: inherit;
        }

        .process-button:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }

        .process-button:disabled {
          opacity: 0.6;
          cursor: not-allowed;
        }

        .spinner {
          width: 20px;
          height: 20px;
          border: 3px solid #ffffff30;
          border-top-color: white;
          border-radius: 50%;
          animation: spin 1s linear infinite;
        }

        @keyframes spin {
          to { transform: rotate(360deg); }
        }

        .app-footer {
          background: #333;
          color: white;
          text-align: center;
          padding: 1rem;
          margin-top: auto;
        }

        .app-footer p {
          margin: 0;
          font-size: 0.9rem;
        }

        @media (max-width: 768px) {
          .app-header h1 {
            font-size: 2rem;
          }
          
          .option-cards {
            grid-template-columns: 1fr;
            gap: 1rem;
          }
          
          .upload-container {
            padding: 1.5rem;
          }
          
          .file-label {
            padding: 2rem;
          }

          .info-panel {
            padding: 1rem;
            margin: 1rem 0;
          }
        }
      `}</style>
    </div>
  );
}

export default App;