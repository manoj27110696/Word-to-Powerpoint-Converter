import React, { useState } from 'react';
import axios from 'axios';
import config from './config';
import './App.css'; // Import the CSS file

function App() {
  const [file, setFile] = useState(null);
  const [downloadLink, setDownloadLink] = useState('');
  const [errorMessage, setErrorMessage] = useState(''); // State for error messages
  const [isLoading, setIsLoading] = useState(false); // State for loading status
  const [validationErrors, setValidationErrors] = useState({}); // State for validation errors

  const handleFileChange = (event) => {
    setFile(event.target.files[0]);
  };

  const validateForm = () => {
    const errors = {};
    if (!file) {
      errors.file = 'File is required';
    } else if (!file.name.endsWith('.docx')) {
      errors.file = 'Only files with extension docx are allowed';
    }
    return errors;
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    const errors = validateForm();
    if (Object.keys(errors).length > 0) {
      setValidationErrors(errors);
      return;
    }

    const formData = new FormData();
    formData.append('docx_file', file);

    setIsLoading(true); // Set loading status to true
    setErrorMessage(''); // Clear any previous error messages
    setDownloadLink(''); // Clear any previous download links
    setValidationErrors({}); // Clear validation errors

    try {
      const response = await axios.post(`${config.backendUrl}/convert`, formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob', // Important for handling binary data
      });

      // Create a download link for the converted PPTX file
      const url = window.URL.createObjectURL(new Blob([response.data]));
      setDownloadLink(url);
    } catch (error) {
      console.error('Error uploading file:', error);
      setErrorMessage('Failed to upload and convert the file. Please try again.'); // Set error message
    } finally {
      setIsLoading(false); // Set loading status to false
    }
  };

  return (
    <div className="App">
      <h1>DOCX to PPTX Converter</h1>
      <form onSubmit={handleSubmit}>
        <input type="file" accept=".docx" onChange={handleFileChange} required />
        {validationErrors.file && <p className="validation-error">{validationErrors.file}</p>} {/* Conditionally render validation error */}
        <button type="submit" disabled={isLoading}>Convert</button>
      </form>
      {isLoading && <p className="loading">Converting...</p>} {/* Conditionally render loading indicator */}
      {errorMessage && <p className="error">{errorMessage}</p>} {/* Conditionally render error message */}
      {downloadLink && (
        <div>
          <a href={downloadLink} download="output.pptx">Download Converted PPTX</a>
        </div>
      )}
    </div>
  );
}

export default App;