import { useState, useCallback } from "react";
import axios from "axios";
import "./MailMergeUploader.css";

const FileInput = ({ label, accept, onChange, disabled, value }) => (
  <div className="file-input-container">
    <label className="file-input-label">{label}</label>
    <div className={`file-input-dropzone ${disabled ? 'disabled' : ''}`}>
      <input
        type="file"
        className="file-input"
        accept={accept}
        onChange={onChange}
        disabled={disabled}
      />
      <div className="file-input-content">
        <svg className="file-input-icon" stroke="currentColor" fill="none" viewBox="0 0 48 48">
          <path d="M28 8H12a4 4 0 00-4 4v20m32-12v8m0 0v8a4 4 0 01-4 4H12a4 4 0 01-4-4v-4m32-4l-3.172-3.172a4 4 0 00-5.656 0L28 28M8 32l9.172-9.172a4 4 0 015.656 0L28 28m0 0l4 4m4-24h8m-4-4v8m-12 4h.02" 
            strokeWidth="2" 
            strokeLinecap="round" 
            strokeLinejoin="round"
          />
        </svg>
        <div className="upload-button">
          {value ? value.name : "Click to upload"}
        </div>
        <p className="upload-text">or drag and drop</p>
        <p className="upload-format">
          {accept === '.docx' ? 'Word document (.docx)' : 'Excel spreadsheet (.xlsx)'}
        </p>
      </div>
    </div>
  </div>
);

export default function MailMergeUploader() {
  const [templateFile, setTemplateFile] = useState(null);
  const [dataFile, setDataFile] = useState(null);
  const [status, setStatus] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [selectedPreview, setSelectedPreview] = useState(null);
  const [zoomLevel, setZoomLevel] = useState(100);
  const [isSending, setIsSending] = useState(false);

  const handleUpload = async (e) => {
    e.preventDefault();
    if (!templateFile || !dataFile) {
      setStatus({ error: "Please upload both template and data files." });
      return;
    }

    const formData = new FormData();
    formData.append("template", templateFile);
    formData.append("datafile", dataFile);

    setIsLoading(true);
    setStatus({ message: "Generating documents..." });

    try {
      const response = await axios.post("/api/upload?skipEmail=true", formData);
      setStatus(response.data);
    } catch (error) {
      console.error("Error during upload:", error);
      setStatus({
        error: error.response?.data?.error || "Failed to upload or process files."
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleSendEmail = async (result) => {
    if (isSending) return;
    
    setIsSending(true);
    try {
      await axios.post(`/api/send-email`, {
        to: result.to,
        files: result.files
      });
      
      setStatus(prevStatus => ({
        ...prevStatus,
        results: prevStatus.results.map(r => 
          r.to === result.to 
            ? { ...r, emailSent: true }
            : r
        )
      }));
    } catch (error) {
      console.error("Error sending email:", error);
      setStatus(prevStatus => ({
        ...prevStatus,
        results: prevStatus.results.map(r => 
          r.to === result.to 
            ? { ...r, emailError: error.response?.data?.error || "Failed to send email" }
            : r
        )
      }));
    } finally {
      setIsSending(false);
    }
  };

  const handlePreview = (filename) => {
    if (filename.endsWith('.pdf')) {
      setSelectedPreview(`/api/preview/${filename}`);
    } else {
      window.open(`/api/files/${filename}`, '_blank');
    }
  };

  const handleDownload = (filename) => {
    window.open(`/api/files/${filename}`, '_blank');
  };

  const handleZoomIn = useCallback(() => {
    setZoomLevel(prev => Math.min(prev + 25, 200));
  }, []);

  const handleZoomOut = useCallback(() => {
    setZoomLevel(prev => Math.max(prev - 25, 50));
  }, []);

  const handleResetZoom = useCallback(() => {
    setZoomLevel(100);
  }, []);

  const handleClosePreview = useCallback(() => {
    setSelectedPreview(null);
    setZoomLevel(100);
  }, []);

  // Close on escape key
  const handleKeyDown = useCallback((e) => {
    if (e.key === 'Escape') {
      handleClosePreview();
    }
  }, [handleClosePreview]);

  // Close on overlay click
  const handleOverlayClick = useCallback((e) => {
    if (e.target === e.currentTarget) {
      handleClosePreview();
    }
  }, [handleClosePreview]);

  return (
    <div className="uploader-container">
      <div className="uploader-wrapper">
        <div className="uploader-card">
          <h1 className="uploader-title">Mail Merge System</h1>
          <p className="uploader-description">Generate and send personalized letters. Upload your template and data files to begin.</p>
          <form onSubmit={handleUpload} className="uploader-form">
            <FileInput
              label="Word Template (.docx)"
              accept=".docx"
              onChange={(e) => setTemplateFile(e.target.files[0])}
              disabled={isLoading}
              value={templateFile}
            />
            <FileInput
              label="Excel Data (.xlsx)"
              accept=".xlsx"
              onChange={(e) => setDataFile(e.target.files[0])}
              disabled={isLoading}
              value={dataFile}
            />
            <button
              type="submit"
              disabled={isLoading || !templateFile || !dataFile}
              className="submit-button"
            >
              {isLoading ? (
                <span className="flex items-center justify-center gap-2">
                  <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none"/>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"/>
                  </svg>
                  Processing...
                </span>
              ) : (
                'Generate & Send Letters'
              )}
            </button>
          </form>
          {status?.error && (
            <div className="status-error">
              <span className="status-error-text">{status.error}</span>
            </div>
          )}
          {status?.message && (
            <div className="status-success">
              <h3 className="status-heading">{status.message}</h3>
              {status.results && (
                <div className="results-container">
                  {status.results.map((result, index) => (
                    <div key={index} className={`result-item ${result.status}`}>
                      <div className="result-content">
                        <span className={`result-text ${result.status}`}>{result.to}: {result.status}</span>
                        {result.error && <span className="result-error">{result.error}</span>}
                        {result.emailError && <span className="result-error">{result.emailError}</span>}
                        {result.status === 'success' && result.files && (
                          <div className="result-buttons">
                            <button onClick={() => handlePreview(result.files.pdf)} className="result-button preview-button">
                              Preview PDF
                            </button>
                            <button onClick={() => handleDownload(result.files.pdf)} className="result-button download-pdf-button">
                              Download PDF
                            </button>
                            <button onClick={() => handleDownload(result.files.docx)} className="result-button download-docx-button">
                              Download DOCX
                            </button>
                            {!result.emailSent && (
                              <button
                                onClick={() => handleSendEmail(result)}
                                className="result-button send-email-button"
                                disabled={isSending}
                              >
                                {isSending ? 'Sending...' : 'Send Email'}
                              </button>
                            )}
                            {result.emailSent && (
                              <span className="email-sent-text">Email Sent ✓</span>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
        {selectedPreview && (
          <div className="preview-modal" onClick={handleOverlayClick} onKeyDown={handleKeyDown} tabIndex={-1}>
            <div className="preview-modal-content">
              <div className="preview-header">
                <h2 className="preview-title">PDF Preview</h2>
                <div className="preview-controls">
                  <div className="zoom-controls">
                    <button onClick={handleZoomOut} className="zoom-button" title="Zoom Out">
                      <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M20 12H4"/>
                      </svg>
                    </button>
                    <span className="zoom-text">{zoomLevel}%</span>
                    <button onClick={handleZoomIn} className="zoom-button" title="Zoom In">
                      <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4"/>
                      </svg>
                    </button>
                    <button onClick={handleResetZoom} className="zoom-button" title="Reset Zoom">
                      <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/>
                      </svg>
                    </button>
                  </div>
                  <button onClick={handleClosePreview} className="close-button" title="Close Preview">
                    <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12"/>
                    </svg>
                  </button>
                </div>
              </div>
              <div className="preview-body">
                <div className="preview-iframe-container">
                  <iframe 
                    src={selectedPreview} 
                    className="preview-iframe" 
                    title="PDF Preview" 
                    style={{
                      transform: `scale(${zoomLevel / 100})`,
                      transformOrigin: 'center center'
                    }}
                  />
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
