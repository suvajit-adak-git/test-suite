import React, { useState } from 'react';
import axios from 'axios';

const TCTraceability = () => {
    const [file, setFile] = useState(null);
    const [fileName, setFileName] = useState('📄 Choose file or drag here');
    const [results, setResults] = useState(null);
    const [error, setError] = useState(null);
    const [loading, setLoading] = useState(false);
    const [activeFilter, setActiveFilter] = useState('failed'); // Default to 'failed'

    const handleFileChange = (e) => {
        if (e.target.files && e.target.files[0]) {
            const selectedFile = e.target.files[0];
            setFile(selectedFile);
            setFileName('✓ ' + selectedFile.name);
            setResults(null);
            setError(null);
            setActiveFilter('failed'); // Reset filter on new file
        }
    };

    const handleDragOver = (e) => {
        e.preventDefault();
        e.currentTarget.style.borderColor = '#4285f4';
        e.currentTarget.style.background = '#e8f0fe';
    };

    const handleDragLeave = (e) => {
        e.preventDefault();
        e.currentTarget.style.borderColor = '#d1d5db';
        e.currentTarget.style.background = '#f8f9fa';
    };

    const handleDrop = (e) => {
        e.preventDefault();
        e.currentTarget.style.borderColor = '#d1d5db';
        e.currentTarget.style.background = '#f8f9fa';

        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            const selectedFile = e.dataTransfer.files[0];
            setFile(selectedFile);
            setFileName('✓ ' + selectedFile.name);
            setResults(null);
            setError(null);
            setActiveFilter('failed');
        }
    };

    const handleSubmit = async () => {
        if (!file) {
            setError('Please select a file');
            return;
        }

        setLoading(true);
        const formData = new FormData();
        formData.append('file', file);

        try {
            const response = await axios.post('http://localhost:8000/api/validate-tc-traceability', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data',
                },
            });
            setResults(response.data);
            setError(null);
            // Ensure filter is set to failed initially
            setActiveFilter('failed');
        } catch (err) {
            setError(err.response ? err.response.data.detail : 'An error occurred');
            setResults(null);
        } finally {
            setLoading(false);
        }
    };

    const getCardStyle = (color, filterType) => ({
        padding: '1rem',
        background: activeFilter === filterType ? `${color}20` : 'white',
        borderRadius: '8px',
        border: activeFilter === filterType ? `2px solid ${color}` : `1px solid ${color}40`,
        textAlign: 'center',
        cursor: 'pointer',
        transition: 'all 0.2s ease',
        transform: activeFilter === filterType ? 'scale(1.02)' : 'scale(1)'
    });

    const filteredResults = results ? results.results.filter(result => {
        const status = result.status ? result.status.toLowerCase() : '';
        if (activeFilter === 'all') return true;
        if (activeFilter === 'passed') return status === 'pass';
        if (activeFilter === 'failed') return status !== 'pass';
        return true;
    }) : [];

    return (
        <div className="card floating" style={{
            background: 'white',
            borderRadius: '20px',
            padding: '3rem',
            boxShadow: '0 10px 40px rgba(0,0,0,0.1)',
            backdropFilter: 'blur(10px)',
            transition: 'all 0.4s ease',
            maxWidth: '900px',
            margin: '0 auto'
        }}>
            <div className="form-group" style={{ marginBottom: '2rem' }}>
                <label className="form-label" style={{
                    display: 'block',
                    fontSize: '1.1rem',
                    fontWeight: '600',
                    marginBottom: '1rem',
                    color: '#333'
                }}>Upload Traceability Matrix</label>
                <div className="file-input-wrapper" style={{
                    position: 'relative',
                    display: 'inline-block',
                    width: '100%'
                }}>
                    <input type="file" id="traceabilityFile" className="file-input"
                        onChange={handleFileChange}
                        style={{ display: 'none' }}
                    />
                    <label htmlFor="traceabilityFile" className={`file-label ${file ? 'has-file' : ''}`}
                        onDragOver={handleDragOver}
                        onDragLeave={handleDragLeave}
                        onDrop={handleDrop}
                        style={{
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            padding: '1.25rem 1.5rem',
                            background: file ? '#e8f0fe' : '#f8f9fa',
                            borderColor: file ? '#4285f4' : '#d1d5db',
                            border: '2px dashed',
                            borderRadius: '12px',
                            cursor: 'pointer',
                            transition: 'all 0.3s ease',
                            fontSize: '0.95rem',
                            color: file ? '#4285f4' : '#6b7280'
                        }}
                    >
                        <span>{fileName}</span>
                        <span className="file-icon" style={{ fontSize: '1.5rem' }}>📊</span>
                    </label>
                </div>
            </div>

            <button className="compare-btn" onClick={handleSubmit} disabled={loading}
                style={{
                    width: '100%',
                    padding: '1.25rem 2rem',
                    background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                    color: 'white',
                    border: 'none',
                    borderRadius: '50px',
                    fontSize: '1.1rem',
                    fontWeight: '600',
                    cursor: 'pointer',
                    transition: 'all 0.4s ease',
                    boxShadow: '0 10px 30px rgba(102, 126, 234, 0.4)',
                    marginBottom: '2rem',
                    opacity: loading ? 0.7 : 1
                }}
            >
                {loading ? 'Validating...' : 'Validate Traceability'}
            </button>

            {error && (
                <div style={{
                    padding: '1rem',
                    background: '#fee2e2',
                    border: '1px solid #fca5a5',
                    borderRadius: '8px',
                    color: '#b91c1c',
                    marginBottom: '2rem'
                }}>
                    {error}
                </div>
            )}

            {results && (
                <div style={{ animation: 'fadeIn 0.5s ease-out' }}>
                    <h3 style={{ marginBottom: '1.5rem', color: '#333' }}>Summary</h3>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '1rem', marginBottom: '2rem' }}>
                        <div style={getCardStyle('#3b82f6', 'all')} onClick={() => setActiveFilter('all')}>
                            <div style={{ fontSize: '0.875rem', color: '#6b7280', marginBottom: '0.5rem' }}>Total Requirements</div>
                            <div style={{ fontSize: '1.5rem', fontWeight: '700', color: '#3b82f6' }}>{results.summary.total_requirements}</div>
                        </div>
                        <div style={getCardStyle('#10b981', 'passed')} onClick={() => setActiveFilter('passed')}>
                            <div style={{ fontSize: '0.875rem', color: '#6b7280', marginBottom: '0.5rem' }}>Passed</div>
                            <div style={{ fontSize: '1.5rem', fontWeight: '700', color: '#10b981' }}>{results.summary.passed}</div>
                        </div>
                        <div style={getCardStyle('#ef4444', 'failed')} onClick={() => setActiveFilter('failed')}>
                            <div style={{ fontSize: '0.875rem', color: '#6b7280', marginBottom: '0.5rem' }}>Failed</div>
                            <div style={{ fontSize: '1.5rem', fontWeight: '700', color: '#ef4444' }}>{results.summary.failed}</div>
                        </div>
                    </div>

                    <h3 style={{ marginBottom: '1rem', color: '#333' }}>
                        {activeFilter === 'all' ? 'All Results' :
                            activeFilter === 'passed' ? 'Passed Requirements' :
                                'Failed Requirements'}
                    </h3>
                    <div style={{ overflowX: 'auto', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem' }}>
                            <thead>
                                <tr style={{ background: '#f9fafb' }}>
                                    <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Requirement ID</th>
                                    <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Expected TCs</th>
                                    <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Found in Sheets</th>
                                    <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Status</th>
                                    <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Error</th>
                                </tr>
                            </thead>
                            <tbody>
                                {filteredResults.length > 0 ? (
                                    filteredResults.map((result, index) => {
                                        const isPass = result.status && result.status.toLowerCase() === 'pass';
                                        return (
                                            <tr key={index} style={{ borderBottom: '1px solid #f3f4f6', background: index % 2 === 0 ? 'white' : '#f9fafb' }}>
                                                <td style={{ padding: '12px', color: '#4b5563' }}>{result.requirement_id}</td>
                                                <td style={{ padding: '12px', color: '#4b5563' }}>{result.expected_tcs.join(', ')}</td>
                                                <td style={{ padding: '12px', color: '#4b5563' }}>{result.found_in_sheets.join(', ')}</td>
                                                <td style={{ padding: '12px' }}>
                                                    <span style={{
                                                        padding: '0.25rem 0.75rem',
                                                        borderRadius: '9999px',
                                                        fontSize: '0.75rem',
                                                        fontWeight: '600',
                                                        background: isPass ? '#d1fae5' : '#fee2e2',
                                                        color: isPass ? '#059669' : '#b91c1c'
                                                    }}>
                                                        {result.status}
                                                    </span>
                                                </td>
                                                <td style={{ padding: '12px', color: '#ef4444', fontSize: '0.85rem' }}>{result.error}</td>
                                            </tr>
                                        );
                                    })
                                ) : (
                                    <tr>
                                        <td colSpan="5" style={{ padding: '2rem', textAlign: 'center', color: '#6b7280' }}>
                                            No results found for this category.
                                        </td>
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>
                </div>
            )}
        </div>
    );
};

export default TCTraceability;
