import React, { useState, useRef, useEffect } from 'react';
import axios from 'axios';

const TCTraceability = () => {
    // Mode state: 'validate' or 'cia'
    const [mode, setMode] = useState('validate');

    // File states
    const [file, setFile] = useState(null);
    const [fileName, setFileName] = useState('📄 Choose file or drag here');
    const [ciaFile, setCiaFile] = useState(null);
    const [ciaFileName, setCiaFileName] = useState('📄 Choose CIA file');

    // Results and UI states
    const [results, setResults] = useState(null);
    const [error, setError] = useState(null);
    const [loading, setLoading] = useState(false);
    const [activeFilter, setActiveFilter] = useState('failed');

    const resultsRef = useRef(null);

    // Auto-scroll to results
    useEffect(() => {
        if (results && resultsRef.current) {
            setTimeout(() => {
                resultsRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 100);
        }
    }, [results]);

    // Reset files when mode changes
    useEffect(() => {
        setFile(null);
        setFileName('📄 Choose file or drag here');
        setCiaFile(null);
        setCiaFileName('📄 Choose CIA file');
        setResults(null);
        setError(null);
    }, [mode]);

    const handleFileChange = (e, isCiaFile = false) => {
        if (e.target.files && e.target.files[0]) {
            const selectedFile = e.target.files[0];
            if (isCiaFile) {
                setCiaFile(selectedFile);
                setCiaFileName('✓ ' + selectedFile.name);
            } else {
                setFile(selectedFile);
                setFileName('✓ ' + selectedFile.name);
            }
            setResults(null);
            setError(null);
            setActiveFilter('failed');
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

    const handleDrop = (e, isCiaFile = false) => {
        e.preventDefault();
        e.currentTarget.style.borderColor = '#d1d5db';
        e.currentTarget.style.background = '#f8f9fa';

        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            const selectedFile = e.dataTransfer.files[0];
            if (isCiaFile) {
                setCiaFile(selectedFile);
                setCiaFileName('✓ ' + selectedFile.name);
            } else {
                setFile(selectedFile);
                setFileName('✓ ' + selectedFile.name);
            }
            setResults(null);
            setError(null);
            setActiveFilter('failed');
        }
    };

    const handleSubmit = async () => {
        if (mode === 'validate') {
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
                setActiveFilter('failed');
            } catch (err) {
                setError(err.response ? err.response.data.detail : 'An error occurred');
                setResults(null);
            } finally {
                setLoading(false);
            }
        } else {
            // CIA Compare mode
            if (!file || !ciaFile) {
                setError('Please select both TC file and CIA file');
                return;
            }

            setLoading(true);

            // Mock CIA comparison results (backend not yet implemented)
            setTimeout(() => {
                const mockCIAResults = {
                    summary: {
                        total_requirements: 10,
                        passed: 7,
                        failed: 3
                    },
                    results: [
                        { tc_requirement: "REQ_001", cia_requirement: "REQ_001", status: "pass" },
                        { tc_requirement: "REQ_002", cia_requirement: "REQ_002", status: "pass" },
                        { tc_requirement: "REQ_003", cia_requirement: "REQ_003", status: "pass" },
                        { tc_requirement: "REQ_004", cia_requirement: "REQ_004", status: "pass" },
                        { tc_requirement: "REQ_005", cia_requirement: "REQ_005", status: "pass" },
                        { tc_requirement: "REQ_006", cia_requirement: "REQ_006", status: "pass" },
                        { tc_requirement: "REQ_007", cia_requirement: "REQ_007", status: "pass" },
                        { tc_requirement: "REQ_008", cia_requirement: "REQ_010", status: "fail" },
                        { tc_requirement: "REQ_009", cia_requirement: "", status: "fail" },
                        { tc_requirement: "", cia_requirement: "REQ_011", status: "fail" }
                    ]
                };
                setResults(mockCIAResults);
                setError(null);
                setActiveFilter('failed');
                setLoading(false);
            }, 1000);
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
        if (mode === 'validate') {
            const status = result.status ? result.status.toLowerCase() : '';
            if (activeFilter === 'all') return true;
            if (activeFilter === 'passed') return status === 'pass';
            if (activeFilter === 'failed') return status !== 'pass';
            return true;
        } else {
            // CIA mode filtering
            const status = result.status ? result.status.toLowerCase() : '';
            if (activeFilter === 'all') return true;
            if (activeFilter === 'passed') return status === 'pass';
            if (activeFilter === 'failed') return status !== 'pass';
            return true;
        }
    }) : [];

    const renderFileUpload = (label, currentFile, currentFileName, isCiaFile = false) => (
        <div className="form-group" style={{ marginBottom: '1.5rem' }}>
            <label className="form-label" style={{
                display: 'block',
                fontSize: '1rem',
                fontWeight: '600',
                marginBottom: '0.75rem',
                color: '#333'
            }}>{label}</label>
            <div className="file-input-wrapper" style={{
                position: 'relative',
                display: 'inline-block',
                width: '100%'
            }}>
                <input
                    type="file"
                    id={isCiaFile ? "ciaFile" : "tcFile"}
                    className="file-input"
                    accept=".xls,.xlsx,.xlsm"
                    onChange={(e) => handleFileChange(e, isCiaFile)}
                    style={{ display: 'none' }}
                />
                <label
                    htmlFor={isCiaFile ? "ciaFile" : "tcFile"}
                    className={`file-label ${currentFile ? 'has-file' : ''}`}
                    onDragOver={handleDragOver}
                    onDragLeave={handleDragLeave}
                    onDrop={(e) => handleDrop(e, isCiaFile)}
                    style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        padding: '1rem 1.25rem',
                        background: currentFile ? '#e8f0fe' : '#f8f9fa',
                        borderColor: currentFile ? '#4285f4' : '#d1d5db',
                        border: '2px dashed',
                        borderRadius: '12px',
                        cursor: 'pointer',
                        transition: 'all 0.3s ease',
                        fontSize: '0.9rem',
                        color: currentFile ? '#4285f4' : '#6b7280'
                    }}
                >
                    <span>{currentFileName}</span>
                    <span className="file-icon" style={{ fontSize: '1.5rem' }}>📊</span>
                </label>
            </div>
        </div>
    );

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
            {/* Mode Toggle Switch */}
            <div style={{ marginBottom: '2rem', display: 'flex', justifyContent: 'center' }}>
                <div style={{
                    display: 'inline-flex',
                    background: '#f3f4f6',
                    borderRadius: '50px',
                    padding: '4px',
                    gap: '4px'
                }}>
                    <button
                        onClick={() => setMode('validate')}
                        style={{
                            padding: '0.75rem 2rem',
                            borderRadius: '50px',
                            border: 'none',
                            background: mode === 'validate' ? 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)' : 'transparent',
                            color: mode === 'validate' ? 'white' : '#6b7280',
                            fontWeight: '600',
                            fontSize: '0.95rem',
                            cursor: 'pointer',
                            transition: 'all 0.3s ease',
                            boxShadow: mode === 'validate' ? '0 4px 12px rgba(102, 126, 234, 0.3)' : 'none'
                        }}
                    >
                        Validate Traceability
                    </button>
                    <button
                        onClick={() => setMode('cia')}
                        style={{
                            padding: '0.75rem 2rem',
                            borderRadius: '50px',
                            border: 'none',
                            background: mode === 'cia' ? 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)' : 'transparent',
                            color: mode === 'cia' ? 'white' : '#6b7280',
                            fontWeight: '600',
                            fontSize: '0.95rem',
                            cursor: 'pointer',
                            transition: 'all 0.3s ease',
                            boxShadow: mode === 'cia' ? '0 4px 12px rgba(102, 126, 234, 0.3)' : 'none'
                        }}
                    >
                        CIA Compare
                    </button>
                </div>
            </div>

            {/* File Upload Section */}
            {mode === 'validate' ? (
                renderFileUpload('Upload Traceability Matrix', file, fileName, false)
            ) : (
                <>
                    {renderFileUpload('TC File', file, fileName, false)}
                    {renderFileUpload('CIA File', ciaFile, ciaFileName, true)}
                </>
            )}

            {/* Submit Button */}
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
                {loading ? 'Processing...' : (mode === 'validate' ? 'Validate Traceability' : 'Compare with CIA')}
            </button>

            {/* Error Display */}
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

            {/* Results Display */}
            {results && (
                <div ref={resultsRef} style={{ animation: 'fadeIn 0.5s ease-out' }}>
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

                    {/* Warnings Section - Only show in Failed view for Validate mode */}
                    {mode === 'validate' && activeFilter === 'failed' && results.summary.warnings && results.summary.warnings.length > 0 && (
                        <div style={{
                            padding: '1rem',
                            background: '#fef3c7',
                            border: '1px solid #fbbf24',
                            borderRadius: '8px',
                            marginBottom: '2rem'
                        }}>
                            <div style={{
                                display: 'flex',
                                alignItems: 'center',
                                gap: '0.5rem',
                                marginBottom: '0.75rem'
                            }}>
                                <span style={{ fontSize: '1.25rem' }}>⚠️</span>
                                <h4 style={{ margin: 0, color: '#92400e', fontSize: '1rem', fontWeight: '600' }}>Warnings</h4>
                            </div>
                            <ul style={{
                                margin: 0,
                                paddingLeft: '1.5rem',
                                color: '#78350f',
                                fontSize: '0.9rem'
                            }}>
                                {results.summary.warnings.map((warning, index) => (
                                    <li key={index} style={{ marginBottom: '0.5rem' }}>{warning}</li>
                                ))}
                            </ul>
                        </div>
                    )}

                    <h3 style={{ marginBottom: '1rem', color: '#333' }}>
                        {activeFilter === 'all' ? 'All Results' :
                            activeFilter === 'passed' ? 'Passed Requirements' :
                                'Failed Requirements'}
                    </h3>

                    {/* Results Table */}
                    <div style={{ overflowX: 'auto', borderRadius: '8px', border: '1px solid #e5e7eb' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem' }}>
                            <thead>
                                <tr style={{ background: '#f9fafb' }}>
                                    {mode === 'validate' ? (
                                        <>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Requirement ID</th>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Expected TCs</th>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Found in Sheets</th>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Status</th>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Error</th>
                                        </>
                                    ) : (
                                        <>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Requirements from TC</th>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Requirements from CIA</th>
                                            <th style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid #e5e7eb', fontWeight: '600', color: '#374151' }}>Status</th>
                                        </>
                                    )}
                                </tr>
                            </thead>
                            <tbody>
                                {filteredResults.length > 0 ? (
                                    filteredResults.map((result, index) => {
                                        const isPass = result.status && result.status.toLowerCase() === 'pass';
                                        return (
                                            <tr key={index} style={{ borderBottom: '1px solid #f3f4f6', background: index % 2 === 0 ? 'white' : '#f9fafb' }}>
                                                {mode === 'validate' ? (
                                                    <>
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
                                                    </>
                                                ) : (
                                                    <>
                                                        <td style={{ padding: '12px', color: '#4b5563' }}>{result.tc_requirement || '-'}</td>
                                                        <td style={{ padding: '12px', color: '#4b5563' }}>{result.cia_requirement || '-'}</td>
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
                                                    </>
                                                )}
                                            </tr>
                                        );
                                    })
                                ) : (
                                    <tr>
                                        <td colSpan={mode === 'validate' ? "5" : "3"} style={{ padding: '2rem', textAlign: 'center', color: '#6b7280' }}>
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
