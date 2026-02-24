import React, { useState } from 'react';
import db from '../services/db';

function Setup({ onSetupComplete }) {
    const [username, setUsername] = useState('admin');
    const [password, setPassword] = useState('');
    const [confirmPassword, setConfirmPassword] = useState('');
    const [error, setError] = useState('');
    const [isSubmitting, setIsSubmitting] = useState(false);

    const handleSubmit = async (e) => {
        e.preventDefault();
        setError('');

        if (!username || !password) {
            setError('Username and password are required.');
            return;
        }

        if (password !== confirmPassword) {
            setError('Passwords do not match.');
            return;
        }

        setIsSubmitting(true);
        try {
            const result = await db.setupAdmin(username, password);
            if (result.success) {
                // Immediately log them in
                onSetupComplete(result.role);
            } else {
                setError(result.error || 'Setup failed.');
            }
        } catch (err) {
            setError('An unexpected error occurred during setup.');
            console.error(err);
        } finally {
            setIsSubmitting(false);
        }
    };

    return (
        <div className="login-screen">
            <div id="particles-js" className="particles-bg"></div>

            <div className="login-container scale-in">
                <div className="system-branding">
                    <img src="assets/images/capas_senior_high_school.jpg" alt="School Logo" className="logo" />
                    <h1>Form 137 System Setup</h1>
                    <p style={{ marginTop: '10px', fontSize: '14px', color: '#fbbf24' }}>
                        Welcome! Please create the Master Administrator account to secure your database.
                    </p>
                </div>

                <div className="login-form-container">
                    <form id="setup-form" onSubmit={handleSubmit}>
                        <div className="input-group">
                            <label htmlFor="setup-username">Admin Username</label>
                            <input
                                type="text"
                                id="setup-username"
                                value={username}
                                onChange={(e) => setUsername(e.target.value)}
                                required
                                autoFocus
                            />
                        </div>

                        <div className="input-group">
                            <label htmlFor="setup-password">Admin Password</label>
                            <div className="password-wrapper">
                                <input
                                    type="password"
                                    id="setup-password"
                                    value={password}
                                    onChange={(e) => setPassword(e.target.value)}
                                    required
                                />
                            </div>
                        </div>

                        <div className="input-group">
                            <label htmlFor="setup-confirm-password">Confirm Password</label>
                            <div className="password-wrapper">
                                <input
                                    type="password"
                                    id="setup-confirm-password"
                                    value={confirmPassword}
                                    onChange={(e) => setConfirmPassword(e.target.value)}
                                    required
                                />
                            </div>
                        </div>

                        {error && (
                            <div style={{ color: '#ef4444', marginBottom: '15px', fontSize: '13px', textAlign: 'center' }}>
                                {error}
                            </div>
                        )}

                        <button
                            type="submit"
                            className="btn btn-primary login-btn"
                            disabled={isSubmitting}
                        >
                            {isSubmitting ? 'CREATING ACCOUNT...' : 'COMPLETE SETUP'}
                        </button>
                    </form>
                </div>
            </div>
        </div>
    );
}

export default Setup;
