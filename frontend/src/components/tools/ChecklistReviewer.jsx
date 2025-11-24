import React, { useState } from 'react';
import { extractHyperlinks, updateBuild } from '../../api/svnApi';

const ChecklistReviewer = () => {
    const [fileName, setFileName] = useState('📋 Choose file or drag here');
    const [file, setFile] = useState(null);
    const [analysisData, setAnalysisData] = useState(null);
    const [loading, setLoading] = useState(false);
    const [newBuild, setNewBuild] = useState('');
    const [updateLoading, setUpdateLoading] = useState(false);

    const handleFileChange = (e) => {
        if (e.target.files && e.target.files[0]) {
            const selectedFile = e.target.files[0];
            setFileName('✓ ' + selectedFile.name);
            setFile(selectedFile);
            setAnalysisData(null); // Reset analysis on new file
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
            setFileName('✓ ' + selectedFile.name);
            setFile(selectedFile);
            setAnalysisData(null);
        }
    };

    const handleAnalyze = async () => {
        if (!file) {
            alert('Please upload a file first.');
            return;
        }

        setLoading(true);
        try {
            const data = await extractHyperlinks(file);
            setAnalysisData(data);
        } catch (error) {
            console.error(error);
            alert(error.message);
        } finally {
            setLoading(false);
        }
    };

    const handleUpdateBuild = async () => {
        if (!file || !newBuild) {
            alert('Please provide both file and new build number.');
            return;
        }

        setUpdateLoading(true);
        try {
            const blob = await updateBuild(file, newBuild);

            // Create download link
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = file.name.replace(/(\.[\w\d_-]+)$/i, `_build_${newBuild}$1`);
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);

            alert('Build updated and file downloaded successfully!');
        } catch (error) {
            console.error(error);
            alert(error.message);
        } finally {
            setUpdateLoading(false);
        }
    };

    const [selectedBuild, setSelectedBuild] = useState(null);

    // Group hyperlinks by build
    const getBuildGroups = () => {
        if (!analysisData || !analysisData.hyperlinks) return {};

        const groups = {};
        analysisData.hyperlinks.forEach(link => {
            if (!groups[link.build]) {
                groups[link.build] = [];
            }
            groups[link.build].push(link);
        });
        return groups;
    };

    const buildGroups = getBuildGroups();
    const uniqueBuilds = Object.keys(buildGroups).sort();

    return (
        <div className="card floating" style={{
            background: 'white',
            borderRadius: '20px',
            padding: '3rem',
            boxShadow: '0 10px 40px rgba(0,0,0,0.1)',
            backdropFilter: 'blur(10px)',
            transition: 'all 0.4s ease',
            maxWidth: '700px',
            margin: '0 auto'
        }}>
            <div className="form-group" style={{ marginBottom: '2rem' }}>
                <label className="form-label" style={{
                    display: 'block',
                    fontSize: '1.1rem',
                    fontWeight: '600',
                    marginBottom: '1rem',
                    color: '#333'
                }}>Upload Review Checklist</label>
                <div className="file-input-wrapper" style={{
                    position: 'relative',
                    display: 'inline-block',
                    width: '100%'
                }}>
                    <input type="file" id="checklistFile" className="file-input" accept=".xls, .xlsx"
                        onChange={handleFileChange}
                        style={{ display: 'none' }}
                    />
                    <label htmlFor="checklistFile" className={`file-label ${file ? 'has-file' : ''}`}
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
                        <span className="file-icon" style={{ fontSize: '1.5rem' }}>📋</span>
                    </label>
                </div>
            </div>

            <button className="compare-btn" onClick={handleAnalyze} disabled={loading}
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
                {loading ? 'Analyzing...' : 'Analyze Checklist'}
            </button>

            {analysisData && (
                <div style={{
                    marginTop: '2rem',
                    padding: '1.5rem',
                    background: '#f0f9ff',
                    borderLeft: '4px solid #4285f4',
                    borderRadius: '8px',
                    color: '#0369a1',
                    fontSize: '1rem',
                    lineHeight: '1.6',
                    marginBottom: '2rem'
                }}>
                    {uniqueBuilds.length === 1 ? (
                        <div>
                            <strong>Result:</strong> The review checklist uses <strong>{uniqueBuilds[0]}</strong> build.
                        </div>
                    ) : (
                        <div>
                            <div style={{ marginBottom: '1rem' }}><strong>Result:</strong> Multiple builds detected:</div>
                            <div style={{ display: 'flex', gap: '1rem', flexWrap: 'wrap' }}>
                                {uniqueBuilds.map(build => (
                                    <button
                                        key={build}
                                        onClick={() => setSelectedBuild(selectedBuild === build ? null : build)}
                                        style={{
                                            padding: '0.75rem 1.5rem',
                                            background: selectedBuild === build ? '#4285f4' : 'white',
                                            color: selectedBuild === build ? 'white' : '#4285f4',
                                            border: '2px solid #4285f4',
                                            borderRadius: '30px',
                                            cursor: 'pointer',
                                            fontWeight: '600',
                                            transition: 'all 0.2s ease',
                                            boxShadow: '0 2px 5px rgba(0,0,0,0.05)'
                                        }}
                                    >
                                        Build {build} ({buildGroups[build].length} files)
                                    </button>
                                ))}
                            </div>

                            {selectedBuild && (
                                <div style={{
                                    marginTop: '1.5rem',
                                    background: 'white',
                                    padding: '1rem',
                                    borderRadius: '8px',
                                    border: '1px solid #e5e7eb',
                                    maxHeight: '300px',
                                    overflowY: 'auto'
                                }}>
                                    <h4 style={{ margin: '0 0 1rem 0', color: '#333', fontSize: '0.95rem' }}>Files using Build {selectedBuild}:</h4>
                                    <ul style={{ margin: 0, paddingLeft: '1.5rem', fontSize: '0.9rem', color: '#555' }}>
                                        {buildGroups[selectedBuild].map((link, idx) => (
                                            <li key={idx} style={{ marginBottom: '0.5rem' }}>
                                                {link.file_name} <span style={{ color: '#9ca3af', fontSize: '0.8rem' }}>({link.sheet_name})</span>
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                            )}
                        </div>
                    )}
                </div>
            )}

            {analysisData && (
                <div style={{
                    marginTop: '2rem',
                    borderTop: '1px solid #eee',
                    paddingTop: '2rem'
                }}>
                    <h3 style={{ marginBottom: '1rem', color: '#333' }}>Update Build Number</h3>
                    <div className="form-group" style={{ marginBottom: '1.5rem' }}>
                        <label style={{ display: 'block', marginBottom: '0.5rem', fontWeight: '600', color: '#555' }}>New Build Number</label>
                        <input
                            type="text"
                            value={newBuild}
                            onChange={(e) => setNewBuild(e.target.value)}
                            placeholder="Enter new build number (e.g., 213)"
                            style={{
                                width: '100%',
                                padding: '1rem',
                                borderRadius: '8px',
                                border: '1px solid #d1d5db',
                                fontSize: '1rem'
                            }}
                        />
                    </div>
                    <button className="compare-btn" onClick={handleUpdateBuild} disabled={updateLoading}
                        style={{
                            width: '100%',
                            padding: '1.25rem 2rem',
                            background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
                            color: 'white',
                            border: 'none',
                            borderRadius: '50px',
                            fontSize: '1.1rem',
                            fontWeight: '600',
                            cursor: 'pointer',
                            transition: 'all 0.4s ease',
                            boxShadow: '0 10px 30px rgba(16, 185, 129, 0.4)',
                            opacity: updateLoading ? 0.7 : 1
                        }}
                    >
                        {updateLoading ? 'Updating...' : 'Update Build & Download'}
                    </button>
                </div>
            )}
        </div>
    );
};

export default ChecklistReviewer;
