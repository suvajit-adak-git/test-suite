import React, { useState } from 'react';
import { useParams, Link } from 'react-router-dom';
import Layout from '../components/Layout';
import { uploadSvnFile, uploadReviewChecklist, compareBoth } from '../api/svnApi';
import ComparisonTable from '../components/ComparisonTable';
import ChecklistReviewer from '../components/tools/ChecklistReviewer';

const Tool = () => {
    const { toolId } = useParams();
    const [svnFileName, setSvnFileName] = useState('📄 Choose file or drag here');
    const [checklistFileName, setChecklistFileName] = useState('📋 Choose file or drag here');
    const [svnFile, setSvnFile] = useState(null);
    const [checklistFile, setChecklistFile] = useState(null);
    const [svnData, setSvnData] = useState(null);
    const [checklistData, setChecklistData] = useState(null);
    const [compareButtonText, setCompareButtonText] = useState('Compare Files');
    const [comparisonData, setComparisonData] = useState(null);

    const handleFileChange = async (e, setFileName, setFile, setData, uploadFn) => {
        if (e.target.files && e.target.files[0]) {
            const file = e.target.files[0];
            setFileName('✓ ' + file.name);
            setFile(file);
            try {
                const response = await uploadFn(file);
                setData(response);
            } catch (error) {
                console.error(error);
                alert(error.message);
            }
        } else {
            setFileName(e.target.id === 'svnFile' ? '📄 Choose file or drag here' : '📋 Choose file or drag here');
            setFile(null);
            setData(null);
        }
    };

    const handleDragOver = (e) => {
        e.preventDefault();
        e.currentTarget.style.borderColor = '#4285f4';
        e.currentTarget.style.background = '#e8f0fe';
    };

    const handleDragLeave = (e, hasFile) => {
        e.preventDefault();
        if (!hasFile) {
            e.currentTarget.style.borderColor = '#d1d5db';
            e.currentTarget.style.background = '#f8f9fa';
        }
    };

    const handleDrop = (e, setFileName, setFile, setData, inputId, uploadFn) => {
        e.preventDefault();
        e.currentTarget.style.borderColor = '#d1d5db';
        e.currentTarget.style.background = '#f8f9fa';

        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            const file = e.dataTransfer.files[0];
            setFileName('✓ ' + file.name);
            setFile(file);
            // Assign files to the hidden input for form submission if needed
            document.getElementById(inputId).files = e.dataTransfer.files;

            // Trigger upload on drop
            uploadFn(file).then(response => {
                setData(response);
            }).catch(error => {
                console.error(error);
                alert(error.message);
            });
        }
    };

    const renderSvnInspector = () => (
        <div className="card floating" style={{
            background: 'white',
            borderRadius: '20px',
            padding: '3rem',
            boxShadow: '0 10px 40px rgba(0,0,0,0.1)',
            backdropFilter: 'blur(10px)',
            transition: 'all 0.4s ease',
            maxWidth: '700px', // Added to reduce card size
            margin: '0 auto' // Added to center the card
        }}>
            <div className="form-group" style={{ marginBottom: '2rem' }}>
                <label className="form-label" style={{
                    display: 'block',
                    fontSize: '1.1rem',
                    fontWeight: '600',
                    marginBottom: '1rem',
                    color: '#333'
                }}>SVN File</label>
                <div className="file-input-wrapper" style={{
                    position: 'relative',
                    display: 'inline-block',
                    width: '100%'
                }}>
                    <input type="file" id="svnFile" className="file-input" accept=".csv, .xls, .xlsx"
                        onChange={(e) => handleFileChange(e, setSvnFileName, setSvnFile, setSvnData, uploadSvnFile)}
                        style={{ display: 'none' }}
                    />
                    <label htmlFor="svnFile" className={`file-label ${svnFile ? 'has-file' : ''}`} id="svnFileLabel"
                        onDragOver={handleDragOver}
                        onDragLeave={(e) => handleDragLeave(e, svnFile)}
                        onDrop={(e) => handleDrop(e, setSvnFileName, setSvnFile, setSvnData, 'svnFile', uploadSvnFile)}
                        style={{
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            padding: '1.25rem 1.5rem',
                            background: svnFile ? '#e8f0fe' : '#f8f9fa',
                            borderColor: svnFile ? '#4285f4' : '#d1d5db',
                            border: '2px dashed',
                            borderRadius: '12px',
                            cursor: 'pointer',
                            transition: 'all 0.3s ease',
                            fontSize: '0.95rem',
                            color: svnFile ? '#4285f4' : '#6b7280'
                        }}
                    >
                        <span id="svnFileName">{svnFileName}</span>
                        <span className="file-icon" style={{ fontSize: '1.5rem' }}>📁</span>
                    </label>
                </div>
            </div>

            <div className="form-group" style={{ marginBottom: '2rem' }}>
                <label className="form-label" style={{
                    display: 'block',
                    fontSize: '1.1rem',
                    fontWeight: '600',
                    marginBottom: '1rem',
                    color: '#333'
                }}>Review Checklist</label>
                <div className="file-input-wrapper" style={{
                    position: 'relative',
                    display: 'inline-block',
                    width: '100%'
                }}>
                    <input type="file" id="checklistFile" className="file-input"
                        onChange={(e) => handleFileChange(e, setChecklistFileName, setChecklistFile, setChecklistData, uploadReviewChecklist)}
                        style={{ display: 'none' }}
                    />
                    <label htmlFor="checklistFile" className={`file-label ${checklistFile ? 'has-file' : ''}`} id="checklistFileLabel"
                        onDragOver={handleDragOver}
                        onDragLeave={(e) => handleDragLeave(e, checklistFile)}
                        onDrop={(e) => handleDrop(e, setChecklistFileName, setChecklistFile, setChecklistData, 'checklistFile', uploadReviewChecklist)}
                        style={{
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            padding: '1.25rem 1.5rem',
                            background: checklistFile ? '#e8f0fe' : '#f8f9fa',
                            borderColor: checklistFile ? '#4285f4' : '#d1d5db',
                            border: '2px dashed',
                            borderRadius: '12px',
                            cursor: 'pointer',
                            transition: 'all 0.3s ease',
                            fontSize: '0.95rem',
                            color: checklistFile ? '#4285f4' : '#6b7280'
                        }}
                    >
                        <span id="checklistFileName">{checklistFileName}</span>
                        <span className="file-icon" style={{ fontSize: '1.5rem' }}>📋</span>
                    </label>
                </div>
            </div>

            <button className="compare-btn" id="compareBtn" onClick={async () => {
                if (!svnData || !checklistData) {
                    alert('⚠️ Please upload both files before comparing.');
                    return;
                }
                try {
                    setCompareButtonText('Comparing...');
                    // Prepare payload: svnData is {status: "ok", data: {...}}, checklistData is {status: "ok", data: [...]}
                    // Backend expects "svn" to be the data object/list, and "checklist" to be the data object/list
                    const payload = {
                        svn: svnData.data,
                        checklist: checklistData
                    };

                    const data = await compareBoth(payload);
                    setComparisonData(data);
                    setCompareButtonText('✓ Comparison Complete!');
                } catch (error) {
                    console.error(error);
                    alert(error.message);
                    setCompareButtonText('Compare Files');
                } finally {
                    setTimeout(() => {
                        setCompareButtonText('Compare Files');
                    }, 2000);
                }
            }}
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
                    marginTop: '2rem',
                    position: 'relative',
                    overflow: 'hidden'
                }}
            >
                {compareButtonText}
            </button>


            <div className="info-text" style={{
                marginTop: '2rem',
                padding: '1.5rem',
                background: '#f0f9ff',
                borderLeft: '4px solid #4285f4',
                borderRadius: '8px',
                color: '#0369a1',
                fontSize: '0.9rem',
                lineHeight: '1.6'
            }}>
                💡 <strong>Tip:</strong> Upload your SVN file and review checklist to perform a comprehensive comparison. The tool will analyze differences and provide detailed insights.
            </div>
            {comparisonData && <ComparisonTable data={comparisonData} />}
        </div>
    );

    const renderDefault = () => (
        <div className="card" style={{ padding: 'var(--space-xl)', minHeight: '400px' }}>
            <p style={{ color: 'var(--text-secondary)' }}>
                This tool is under construction.
            </p>
        </div>
    );

    return (
        <Layout>
            <section className="tools-section">
                <div style={{ padding: '0' }}>
                    <Link to="/" style={{ color: 'var(--text-secondary)', marginBottom: 'var(--space-md)', display: 'inline-block' }}>
                        ← Back to Index
                    </Link>
                    <h1 style={{ fontSize: '2.5rem', marginBottom: 'var(--space-md)' }}>
                        Tool: <span className="text-gradient" style={{ textTransform: 'capitalize' }}>{toolId}</span>
                    </h1>
                    {toolId === 'svn' ? renderSvnInspector() :
                        toolId === 'checklist' ? <ChecklistReviewer /> :
                            renderDefault()}
                </div>
            </section>
        </Layout>
    );
};

export default Tool;

