import React, { useState, useEffect, useRef, useCallback } from 'react';
import db from '../services/db';

function Login({ onLogin }) {
    const [username, setUsername] = useState('');
    const [password, setPassword] = useState('');
    const [showPassword, setShowPassword] = useState(false);
    const [error, setError] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    const [alertMessage, setAlertMessage] = useState('');

    const canvasRef = useRef(null);
    const containerRef = useRef(null);
    const formWrapperRef = useRef(null);
    const particlesRef = useRef([]);
    const animationRef = useRef(null);

    // Particle system
    useEffect(() => {
        const canvas = canvasRef.current;
        const container = containerRef.current;
        if (!canvas || !container) return;

        const ctx = canvas.getContext('2d');
        let width, height;

        const resize = () => {
            width = container.clientWidth;
            height = container.clientHeight;
            canvas.width = width;
            canvas.height = height;
        };

        class Particle {
            constructor(x, y, type = 'normal') {
                this.x = x ?? Math.random() * width;
                this.y = y ?? Math.random() * height;
                this.vx = (Math.random() - 0.5) * 2;
                this.vy = (Math.random() - 0.5) * 2;
                this.size = Math.random() * 3 + 2;
                this.color = `rgba(255, 255, 255, ${Math.random() * 0.3 + 0.1})`;
                this.type = type;
                if (type === 'burst') {
                    this.vx *= 3;
                    this.vy *= 3;
                    this.life = 1.0;
                    this.color = `rgba(100, 200, 255, 0.8)`;
                }
            }

            update() {
                this.x += this.vx;
                this.y += this.vy;
                if (this.x < 0 || this.x > width) this.vx *= -1;
                if (this.y < 0 || this.y > height) this.vy *= -1;
                if (this.type === 'burst') {
                    this.life -= 0.02;
                    this.color = `rgba(100, 200, 255, ${this.life})`;
                }
            }

            draw() {
                if (this.type === 'burst' && this.life <= 0) return;
                ctx.beginPath();
                ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
                ctx.fillStyle = this.color;
                ctx.fill();
            }
        }

        resize();
        window.addEventListener('resize', resize);

        // Initial particles
        particlesRef.current = [];
        for (let i = 0; i < 30; i++) {
            particlesRef.current.push(new Particle());
        }

        const animate = () => {
            if (document.hidden) {
                animationRef.current = null;
                return;
            }
            ctx.clearRect(0, 0, width, height);

            const particles = particlesRef.current;

            // Batch draw normal particles
            ctx.beginPath();
            ctx.fillStyle = 'rgba(255, 255, 255, 0.2)'; // Use a constant for normal particles

            for (let i = particles.length - 1; i >= 0; i--) {
                const p = particles[i];
                p.update();

                if (p.type === 'normal') {
                    ctx.moveTo(p.x + p.size, p.y);
                    ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
                } else {
                    // Draw bursts immediately (they have unique colors/alphas)
                    p.draw();
                    if (p.life <= 0) {
                        particles.splice(i, 1);
                    }
                }
            }
            ctx.fill();

            animationRef.current = requestAnimationFrame(animate);
        };
        animate();

        const handleClick = (e) => {
            const x = e.nativeEvent.offsetX;
            const y = e.nativeEvent.offsetY;
            const particles = particlesRef.current;
            for (let i = 0; i < 10; i++) {
                particles.push(new Particle(x, y, 'burst'));
            }
        };

        const handleVisibility = () => {
            if (!document.hidden && !animationRef.current) {
                animate();
            }
        };

        container.addEventListener('mousedown', handleClick);
        document.addEventListener('visibilitychange', handleVisibility);

        return () => {
            window.removeEventListener('resize', resize);
            container.removeEventListener('mousedown', handleClick);
            document.removeEventListener('visibilitychange', handleVisibility);
            if (animationRef.current) {
                cancelAnimationFrame(animationRef.current);
            }
        };
    }, []);

    const handleSubmit = async (e) => {
        e?.preventDefault();

        if (!username.trim() || !password.trim()) {
            setAlertMessage('Please enter both username and password.');
            return;
        }

        setError('');
        setIsLoading(true);

        try {
            const result = await db.login(username.trim(), password.trim());

            if (result.success) {
                onLogin(result.role);
            } else {
                setError('Invalid Username or Password');
                setAlertMessage('Login Failed: Invalid Username or Password');

                // Shake animation
                if (formWrapperRef.current) {
                    formWrapperRef.current.style.transform = 'translateX(10px)';
                    setTimeout(() => {
                        formWrapperRef.current.style.transform = 'translateX(-10px)';
                    }, 100);
                    setTimeout(() => {
                        formWrapperRef.current.style.transform = 'translateX(0)';
                    }, 200);
                }
            }
        } catch (err) {
            console.error(err);
            setError('An error occurred during login.');
            setAlertMessage('System Error: ' + err.message);
        } finally {
            setIsLoading(false);
        }
    };

    const handleKeyPress = (e) => {
        if (e.key === 'Enter') handleSubmit();
    };

    const closeAlert = useCallback(() => {
        setAlertMessage('');
    }, []);

    useEffect(() => {
        const handleKeydown = (e) => {
            if (alertMessage && (e.key === 'Enter' || e.key === 'Escape')) {
                closeAlert();
            }
        };
        document.addEventListener('keydown', handleKeydown);
        return () => document.removeEventListener('keydown', handleKeydown);
    }, [alertMessage, closeAlert]);

    return (
        <div id="login-view">
            <div className="login-container">
                {/* Left Side: Brand/Image */}
                <div className="login-brand-side" id="brand-side" ref={containerRef}>
                    <canvas id="particle-canvas" ref={canvasRef}></canvas>
                    <div className="brand-content">
                        <img
                            src="/assets/images/capas_senior_high_school.jpg"
                            alt="School Logo"
                            className="brand-logo"
                        />
                        <h1>Capas Senior High School</h1>
                        <p>Student Permanent Record System</p>
                    </div>
                </div>

                {/* Right Side: Login Form */}
                <div className="login-form-side">
                    <div className="form-wrapper" ref={formWrapperRef}>
                        <div className="form-header">
                            <h2>Welcome Back</h2>
                            <p>Please sign in to your account</p>
                        </div>

                        {error && (
                            <div
                                id="login-error"
                                style={{
                                    color: '#ef4444',
                                    background: 'rgba(239, 68, 68, 0.1)',
                                    padding: '10px',
                                    borderRadius: '8px',
                                    marginBottom: '20px',
                                    fontSize: '14px',
                                    textAlign: 'center',
                                    border: '1px solid rgba(239, 68, 68, 0.2)'
                                }}
                            >
                                {error}
                            </div>
                        )}

                        <div className="input-group">
                            <label htmlFor="u-input">Username</label>
                            <input
                                type="text"
                                id="u-input"
                                className="login-input"
                                placeholder="Enter your username"
                                value={username}
                                onChange={(e) => setUsername(e.target.value)}
                                disabled={isLoading}
                                autoComplete="username"
                            />
                        </div>

                        <div className="input-group">
                            <label htmlFor="p-input">Password</label>
                            <div style={{ position: 'relative' }}>
                                <input
                                    type={showPassword ? 'text' : 'password'}
                                    id="p-input"
                                    className="login-input"
                                    placeholder="Enter your password"
                                    style={{ paddingRight: '40px' }}
                                    value={password}
                                    onChange={(e) => setPassword(e.target.value)}
                                    onKeyPress={handleKeyPress}
                                    disabled={isLoading}
                                    autoComplete="current-password"
                                />
                                <span
                                    id="toggle-password"
                                    onClick={() => setShowPassword(!showPassword)}
                                    style={{
                                        position: 'absolute',
                                        right: '10px',
                                        top: '50%',
                                        transform: 'translateY(-50%)',
                                        cursor: 'pointer',
                                        color: showPassword ? '#3b82f6' : '#cbd5e1'
                                    }}
                                >
                                    {showPassword ? (
                                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                            <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path>
                                            <line x1="1" y1="1" x2="23" y2="23"></line>
                                        </svg>
                                    ) : (
                                        <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                                            <circle cx="12" cy="12" r="3"></circle>
                                        </svg>
                                    )}
                                </span>
                            </div>
                        </div>

                        <button
                            className="btn-primary"
                            id="login-btn"
                            onClick={handleSubmit}
                            disabled={isLoading}
                        >
                            {isLoading ? 'Signing In...' : 'Sign In'}
                        </button>

                        <div className="form-footer">
                            <p>Form 137 / SF10 System v1.0</p>
                        </div>
                    </div>
                </div>
            </div>

            {/* Custom Alert Modal */}
            {alertMessage && (
                <div id="custom-alert-modal" className="modal-overlay" style={{ display: 'flex' }}>
                    <div className="modal-content premium-modal center-text">
                        <div className="modal-icon warning">
                            <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#fbbf24" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginBottom: '15px' }}><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>
                        </div>
                        <h3 id="custom-alert-message">{alertMessage}</h3>
                        <button
                            className="btn-primary full-width"
                            id="custom-alert-ok"
                            onClick={closeAlert}
                            autoFocus
                        >
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><polyline points="20 6 9 17 4 12"></polyline></svg> OK
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
}

export default Login;
