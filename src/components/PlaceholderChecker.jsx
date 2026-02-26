import React, { useState } from 'react';
// db service available if needed in future

function getIpcRenderer() {
    try {
        if (window.ipcRenderer) return window.ipcRenderer;
        if (window.electron && window.electron.ipcRenderer) return window.electron.ipcRenderer;
        if (window.require) {
            const electron = window.require('electron');
            if (electron && electron.ipcRenderer) return electron.ipcRenderer;
        }
    } catch (e) {
        // Not running in Electron.
    }
    return null;
}

function PlaceholderChecker({ onClose }) {
    const [status, setStatus] = useState('ready'); // ready, scanning, done
    const [placeholders, setPlaceholders] = useState([]);
    const [error, setError] = useState(null);

    const scanTemplate = async () => {
        const ipcRenderer = getIpcRenderer();
        if (!ipcRenderer) {
            setError('Template scanning is available only in the desktop app.');
            setStatus('ready');
            return;
        }

        setStatus('scanning');
        setError(null);
        try {
            const result = await ipcRenderer.invoke('scan-placeholders');
            if (result.success) {
                setPlaceholders(result.placeholders.sort());
                setStatus('done');
            } else {
                setError(result.error || 'Failed to scan template');
                setStatus('ready');
            }
        } catch (e) {
            setError(e.message);
            setStatus('ready');
        }
    };

    return (
        <div className="modal-overlay" style={{ display: 'flex' }}>
            <div className="modal-content premium-modal" style={{ maxWidth: '600px', width: '90%' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                    <h2 style={{ margin: 0 }}>Template Placeholder Checker</h2>
                    <button onClick={onClose} style={{ background: 'none', border: 'none', color: '#888', cursor: 'pointer', fontSize: '20px' }}>&times;</button>
                </div>

                <p style={{ color: '#cbd5e1', marginBottom: '20px', fontSize: '14px' }}>
                    This tool scans <strong>Form 137-SHS-BLANK.xlsx</strong> for <code>%(field)</code> tags to ensure your custom template is set up correctly.
                </p>

                {status === 'ready' && (
                    <button className="btn-primary full-width" onClick={scanTemplate}>
                        Scan Template Now
                    </button>
                )}

                {status === 'scanning' && (
                    <div style={{ textAlign: 'center', padding: '20px' }}>
                        <div className="spinner" style={{ margin: '0 auto 10px auto' }}></div>
                        <p>Scanning Excel sheets...</p>
                    </div>
                )}

                {status === 'done' && (
                    <div className="placeholder-results">
                        <div style={{ maxHeight: '300px', overflowY: 'auto', background: 'rgba(0,0,0,0.2)', padding: '15px', borderRadius: '8px' }}>
                            <h4 style={{ marginTop: 0 }}>Found {placeholders.length} Placeholders:</h4>
                            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                                {placeholders.map(p => (
                                    <div key={p} style={{ fontSize: '12px', fontFamily: 'monospace', color: '#4ade80', background: 'rgba(74, 222, 128, 0.1)', padding: '4px 8px', borderRadius: '4px' }}>
                                        {p}
                                    </div>
                                ))}
                            </div>
                            {placeholders.length === 0 && <p style={{ color: '#f87171' }}>No placeholders found! Use %(lname), %(fname), etc. in your Excel cells.</p>}
                        </div>
                        <button className="btn-primary full-width" style={{ marginTop: '20px' }} onClick={scanTemplate}>
                            Re-scan Template
                        </button>
                    </div>
                )}

                {error && (
                    <div style={{ background: 'rgba(248, 113, 113, 0.1)', color: '#f87171', padding: '10px', borderRadius: '8px', marginBottom: '15px' }}>
                        <strong>Error:</strong> {error}
                    </div>
                )}

                <button className="btn-secondary full-width" style={{ marginTop: '10px' }} onClick={onClose}>
                    Close
                </button>
            </div>
        </div>
    );
}

export default PlaceholderChecker;
