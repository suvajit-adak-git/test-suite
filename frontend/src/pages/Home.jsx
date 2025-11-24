
import React from 'react';
import Layout from '../components/Layout';
import Card from '../components/Card';
import ParticleBackground from '../components/ParticleBackground';
import '../styles/Home.css';
import svnImage from '../assets/svn-inspector.svg';

const Home = () => {
    const tools = [
        {
            id: 'svn',
            title: 'SVN Inspector',
            description: 'Upload and analyze SVN reports',
            route: '/svn-inspector',
            previewText: 'SVN',
            image: svnImage
        },
        {
            id: 'checklist',
            title: 'Checklist Reviewer',
            description: 'Process review checklists',
            route: '/checklist',
            previewText: 'CHECKLIST'
        },
        {
            id: 'compare',
            title: 'Comparator',
            description: 'Compare SVN vs Checklist data',
            route: '/compare',
            previewText: 'COMPARE'
        },
        {
            id: 'tc-traceability',
            title: 'TC Traceability',
            description: 'Validate TC traceability',
            route: '/tc-traceability',
            previewText: 'TC'
        }
    ];

    return (
        <Layout>
            <ParticleBackground />
            <section className="hero-section">
                <h1 className="hero-headline">
                    Experience liftoff<br />
                    <span className="hero-subheadline">with the next-generation IDE</span>
                </h1>
                <div className="hero-actions">
                    <button className="lift-off-btn">
                        <span className="apple-icon"></span> Download for MacOS
                    </button>
                    <button
                        className="btn-hero-secondary"
                        onClick={() => document.getElementById('tools-section')?.scrollIntoView({ behavior: 'smooth' })}
                    >
                        Explore use cases
                    </button>
                </div>
            </section>

            <section className="tools-section" id="tools-section">
                <h2 className="section-title">Available Tools</h2>
                <div className="tools-grid">
                    {tools.map(tool => (
                        <Card
                            key={tool.id}
                            title={tool.title}
                            description={tool.description}
                            route={`/${tool.id}`}
                            previewText={tool.previewText}
                            image={tool.image}
                        />
                    ))}
                </div>
            </section>
        </Layout>
    );
};

export default Home;
