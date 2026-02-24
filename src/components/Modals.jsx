import React, { useState, useEffect, useRef } from 'react';

function Modals({
    alertModal,
    onAlertClose,
    promptModal,
    onPromptSubmit,
    onPromptCancel,
    confirmModal,
    onConfirmYes,
    onConfirmNo
}) {
    const [promptValue, setPromptValue] = useState('');
    const promptInputRef = useRef(null);
    const alertOkRef = useRef(null);
    const confirmNoRef = useRef(null);

    // Focus management
    useEffect(() => {
        if (promptModal.show && promptInputRef.current) {
            setTimeout(() => promptInputRef.current.focus(), 50);
        }
    }, [promptModal.show]);

    useEffect(() => {
        if (alertModal.show && alertOkRef.current) {
            setTimeout(() => alertOkRef.current.focus(), 50);
        }
    }, [alertModal.show]);

    useEffect(() => {
        if (confirmModal.show && confirmNoRef.current) {
            setTimeout(() => confirmNoRef.current.focus(), 50);
        }
    }, [confirmModal.show]);

    // Reset prompt value when modal opens
    useEffect(() => {
        if (promptModal.show) {
            setPromptValue('');
        }
    }, [promptModal.show]);

    // Keyboard handlers
    useEffect(() => {
        const handleKeydown = (e) => {
            if (alertModal.show && (e.key === 'Enter' || e.key === 'Escape')) {
                onAlertClose();
            }
            if (confirmModal.show && e.key === 'Escape') {
                onConfirmNo();
            }
        };
        document.addEventListener('keydown', handleKeydown);
        return () => document.removeEventListener('keydown', handleKeydown);
    }, [alertModal.show, confirmModal.show, onAlertClose, onConfirmNo]);

    const handlePromptKeydown = (e) => {
        if (e.key === 'Enter') {
            onPromptSubmit(promptValue.trim() || null);
        }
        if (e.key === 'Escape') {
            onPromptCancel();
        }
    };

    return (
        <>
            {/* Alert Modal */}
            {alertModal.show && (
                <div id="custom-alert-modal" className="modal-overlay" style={{ display: 'flex' }}>
                    <div className="modal-content premium-modal center-text">
                        <div className="modal-icon warning">
                            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#fbbf24" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginBottom: '15px' }}><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>
                        </div>
                        <h3 id="custom-alert-message">{alertModal.message}</h3>
                        <button
                            ref={alertOkRef}
                            className="btn-primary full-width"
                            onClick={onAlertClose}
                        >
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><polyline points="20 6 9 17 4 12"></polyline></svg> OK
                        </button>
                    </div>
                </div>
            )}

            {/* Prompt Modal */}
            {promptModal.show && (
                <div id="custom-prompt-modal" className="modal-overlay" style={{ display: 'flex' }}>
                    <div className="modal-content premium-modal">
                        <h3 id="custom-prompt-title">{promptModal.title}</h3>
                        <input
                            ref={promptInputRef}
                            type="text"
                            id="custom-prompt-input"
                            className="premium-input"
                            value={promptValue}
                            onChange={(e) => setPromptValue(e.target.value)}
                            onKeyDown={handlePromptKeydown}
                        />
                        <div className="modal-actions">
                            <button
                                className="btn-primary"
                                onClick={() => onPromptSubmit(promptValue.trim() || null)}
                            >
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><polyline points="20 6 9 17 4 12"></polyline></svg> OK
                            </button>
                            <button className="btn-secondary" onClick={onPromptCancel}>
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg> Cancel
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Confirm Modal */}
            {confirmModal.show && (
                <div id="custom-confirm-modal" className="modal-overlay" style={{ display: 'flex' }}>
                    <div className="modal-content premium-modal">
                        <h3 id="custom-confirm-message">{confirmModal.message}</h3>
                        <div className="modal-actions">
                            <button
                                className="btn-primary"
                                style={{ background: 'linear-gradient(135deg, #ef4444 0%, #dc2626 100%)' }}
                                onClick={onConfirmYes}
                            >
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><polyline points="20 6 9 17 4 12"></polyline></svg> Yes, Delete
                            </button>
                            <button ref={confirmNoRef} className="btn-secondary" onClick={onConfirmNo}>
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg> Cancel
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </>
    );
}

export default Modals;
