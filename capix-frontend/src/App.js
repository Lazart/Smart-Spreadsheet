import React, { useState } from 'react';
import { useDropzone } from 'react-dropzone';
import axios from 'axios';
import './App.css';
function App() {
  const [question, setQuestion] = useState('');
  const [fileMissing, setFileMissing] = useState(true);
  const [answers, setAnswers] = useState([]);
  const [loading, setLoading] = useState(false);

  const onDrop = (acceptedFiles) => {
    setLoading(true);
    const file = acceptedFiles[0];
    const formData = new FormData();
    formData.append('file', file);

    axios.post('http://localhost:8000/upload/', formData, {
      headers: {
        'Content-Type': 'multipart/form-data'
      }
    }).then(response => {
      console.log('File uploaded successfully');
      setLoading(false);
      setFileMissing(false)
    }).catch(error => {
      console.error('Error uploading file:', error);
      setLoading(false);
      setFileMissing(true)
    });
  };

  const { getRootProps, getInputProps } = useDropzone({ onDrop });

  const handleQuestionChange = (e) => {
    setQuestion(e.target.value);
  };

  const handleKeyPress = (e) => {
    if (e.key === 'Enter') {
      submitQuestion();
    }
  };

  const submitQuestion = () => {
    setLoading(true);
    axios.get(`http://localhost:8000/ask/?question=${encodeURIComponent(question)}`)
      .then(response => {
        setAnswers(prevAnswers => [...prevAnswers, { question: question, answer: response.data.answer }]);
        setLoading(false);
        setQuestion('');
      })
      .catch(error => {
        console.error('Error fetching answer:', error);
        if (error.response.code === 400) {
          setFileMissing(true)
        }
        setLoading(false);
        setQuestion('');
      });
  };

  return (
    <div className="App" style={{ background: loading ? 'dimgray' : 'white' }}>
      <div {...getRootProps()} style={{ border: '2px dashed gray', padding: '20px', width: '300px', margin: '20px auto' }}>
        <input disabled={loading} {...getInputProps()} />
        <p>Drag 'n' drop an XLSX file here, or click to select file</p>
      </div>

      {loading &&
        <div class="loading-container">
          <div class="loading"></div>
          <div id="loading-text">capix wants to hire lazar</div>
        </div>}

      {fileMissing && <div className="answer-container">
        <div className="answer">
          <p><strong>first upload excel file for context</strong></p>
        </div>
      </div>}
      <div className="answer-container">
        {answers.map((item, index) => (
          <div key={index} className="answer">
            <p><strong>Q:</strong> {item.question}</p>
            <p><strong>A:</strong> {item.answer}</p>
          </div>
        ))}
      </div>
      <div className="input-container">
        <input disabled={loading || fileMissing} type="text" value={question} onChange={handleQuestionChange} onKeyDown={handleKeyPress} placeholder="Ask a question" />
        <button disabled={loading || fileMissing} onClick={submitQuestion}>Submit</button>
      </div>
    </div>
  );
}

export default App;
