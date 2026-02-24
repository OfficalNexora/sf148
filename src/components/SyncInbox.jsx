import React, { useState } from 'react';
import syncService from '../services/syncService';
import db from '../services/db';

function SyncInbox({ requests, onMerge, onDismiss }) {
    const [mergingId, setMergingId] = useState(null);
    const [dismissingId, setDismissingId] = useState(null);

    if (!requests || requests.length === 0) {
        return (
            <div style={{ padding: '20px', textAlign: 'center', color: '#64748b' }}>
                <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" style={{ marginBottom: '10px', display: 'block', margin: '0 auto 10px' }}>
                    <polyline points="9 11 12 14 22 4"></polyline>
                    <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path>
                </svg>
                <p style={{ margin: 0, fontSize: '13px' }}>No pending sync requests</p>
            </div>
        );
    }

    const handleMerge = async (req) => {
        setMergingId(req.id);
        try {
            const syncData = syncService.decodeSyncRequest(req);
            if (!syncData) {
                alert('Failed to decode sync data.');
                return;
            }
            const result = await db.importSyncData(syncData);
            if (!result.success) {
                alert('Import failed: ' + (result.error || 'Unknown error'));
                return;
            }
            await syncService.markAsMerged(req.id);
            if (onMerge) onMerge(req);
            alert(`Successfully merged student records from ${req.teacherName}.`);
        } catch (err) {
            alert('Merge failed: ' + err.message);
        } finally {
            setMergingId(null);
        }
    };

    const handleDismiss = async (req) => {
        setDismissingId(req.id);
        try {
            await syncService.dismissSync(req.id);
            if (onDismiss) onDismiss(req);
        } finally {
            setDismissingId(null);
        }
    };

    const formatTime = (date) => {
        if (!date) return '';
        const now = new Date();
        const diff = Math.floor((now - date) / 1000);
        if (diff < 60) return `${diff}s ago`;
        if (diff < 3600) return `${Math.floor(diff / 60)}m ago`;
        return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    };

    const formatSize = (bytes) => {
        if (!bytes) return '';
        if (bytes < 1024) return `${bytes}B`;
        return `${Math.round(bytes / 1024)}KB`;
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '8px' }}>
            {requests.map(req => (
                <div
                    key={req.id}
                    style={{
                        background: 'rgba(255,255,255,0.05)',
                        border: '1px solid rgba(255,255,255,0.1)',
                        borderRadius: '8px',
                        padding: '12px',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '8px'
                    }}
                >
                    {/* Header */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div>
                            <div style={{ fontWeight: 600, fontSize: '13px', color: 'white' }}>
                                {req.teacherName}
                            </div>
                            <div style={{ fontSize: '11px', color: '#64748b', marginTop: '2px' }}>
                                {formatTime(req.createdAt)} · {formatSize(req.sizeBytes)}
                            </div>
                        </div>
                        <span style={{
                            background: 'rgba(251,191,36,0.15)',
                            color: '#fbbf24',
                            borderRadius: '4px',
                            padding: '2px 6px',
                            fontSize: '10px',
                            fontWeight: 600,
                            letterSpacing: '0.5px'
                        }}>
                            PENDING
                        </span>
                    </div>

                    {/* Actions */}
                    <div style={{ display: 'flex', gap: '6px' }}>
                        <button
                            onClick={() => handleMerge(req)}
                            disabled={mergingId === req.id}
                            style={{
                                flex: 1,
                                padding: '6px',
                                borderRadius: '6px',
                                border: 'none',
                                background: 'linear-gradient(135deg, #8b5cf6, #7c3aed)',
                                color: 'white',
                                fontSize: '12px',
                                fontWeight: 600,
                                cursor: mergingId === req.id ? 'not-allowed' : 'pointer',
                                opacity: mergingId === req.id ? 0.7 : 1,
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                gap: '4px'
                            }}
                        >
                            {mergingId === req.id ? (
                                'Merging...'
                            ) : (
                                <>
                                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                                        <polyline points="20 6 9 17 4 12"></polyline>
                                    </svg>
                                    Auto-Merge
                                </>
                            )}
                        </button>
                        <button
                            onClick={() => handleDismiss(req)}
                            disabled={dismissingId === req.id}
                            style={{
                                padding: '6px 10px',
                                borderRadius: '6px',
                                border: '1px solid rgba(255,255,255,0.1)',
                                background: 'transparent',
                                color: '#94a3b8',
                                fontSize: '12px',
                                cursor: dismissingId === req.id ? 'not-allowed' : 'pointer',
                                opacity: dismissingId === req.id ? 0.7 : 1
                            }}
                        >
                            Dismiss
                        </button>
                    </div>
                </div>
            ))}
        </div>
    );
}

export default SyncInbox;
