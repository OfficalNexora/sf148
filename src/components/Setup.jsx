import React, { useState, useEffect, useRef, useCallback } from 'react';
import db from '../services/db';

function Setup({ onSetupComplete }) {
    const [username, setUsername] = useState('admin');
    const [password, setPassword] = useState('');
    const [confirmPassword, setConfirmPassword] = useState('');
    const [error, setError] = useState('');
    const [isSubmitting, setIsSubmitting] = useState(false);

    const canvasRef = useRef(null);
    const containerRef = useRef(null);
    const formWrapperRef = useRef(null);
    const particlesRef = useRef([]);
    const animationRef = useRef(null);

    // Particle system (Sync with Login.jsx)
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

            // Loop backwards for safe splicing
            const particles = particlesRef.current;
            for (let i = particles.length - 1; i >= 0; i--) {
                const p = particles[i];
                p.update();
                p.draw();
                if (p.type === 'burst' && p.life <= 0) {
                    particles.splice(i, 1);
                }
            }
            animationRef.current = requestAnimationFrame(animate);
        };
        animate();

        const handleClick = (e) => {
            const x = e.nativeEvent.offsetX;
            const y = e.nativeEvent.offsetY;
            for (let i = 0; i < 10; i++) {
                particlesRef.current.push(new Particle(x, y, 'burst'));
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
        setError('');

        if (!username.trim() || !password.trim()) {
            setError('Username and password are required.');
            return;
        }

        if (password !== confirmPassword) {
            setError('Passwords do not match.');
            return;
        }

        setIsSubmitting(true);
        try {
            const result = await db.setupAdmin(username.trim(), password);
            if (result.success) {
                onSetupComplete(result.role);
            } else {
                setError(result.error || 'Setup failed.');
                // Shake effect
                if (formWrapperRef.current) {
                    formWrapperRef.current.style.transform = 'translateX(10px)';
                    setTimeout(() => formWrapperRef.current.style.transform = 'translateX(-10px)', 100);
                    setTimeout(() => formWrapperRef.current.style.transform = 'translateX(0)', 200);
                }
            }
        } catch (err) {
            setError('An unexpected error occurred during setup.');
            console.error(err);
        } finally {
            setIsSubmitting(false);
        }
    };

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

                {/* Right Side: Setup Form */}
                <div className="login-form-side">
                    <div className="form-wrapper" ref={formWrapperRef}>
                        <div className="form-header">
                            <h2>System Setup</h2>
                            <p style={{ marginTop: '10px', fontSize: '14px', color: '#fbbf24' }}>
                                Create the Master Administrator account to secure your database.
                            </p>
                        </div>

                        <form onSubmit={handleSubmit}>
                            <div className="input-group">
                                <label htmlFor="setup-username">Admin Username</label>
                                <input
                                    type="text"
                                    id="setup-username"
                                    className="login-input"
                                    value={username}
                                    onChange={(e) => setUsername(e.target.value)}
                                    required
                                    autoFocus
                                    disabled={isSubmitting}
                                    autoComplete="username"
                                />
                            </div>

                            <div className="input-group">
                                <label htmlFor="setup-password">Admin Password</label>
                                <input
                                    type="password"
                                    id="setup-password"
                                    className="login-input"
                                    value={password}
                                    onChange={(e) => setPassword(e.target.value)}
                                    required
                                    disabled={isSubmitting}
                                    autoComplete="new-password"
                                />
                            </div>

                            <div className="input-group">
                                <label htmlFor="setup-confirm-password">Confirm Password</label>
                                <input
                                    type="password"
                                    id="setup-confirm-password"
                                    className="login-input"
                                    value={confirmPassword}
                                    onChange={(e) => setConfirmPassword(e.target.value)}
                                    required
                                    disabled={isSubmitting}
                                    autoComplete="new-password"
                                />
                            </div>

                            {error && (
                                <div className="error-message">
                                    {error}
                                </div>
                            )}

                            <button
                                type="submit"
                                className="btn-primary"
                                id="login-btn"
                                disabled={isSubmitting}
                                style={{ width: '100%' }}
                            >
                                {isSubmitting ? 'CREATING ACCOUNT...' : 'COMPLETE SETUP'}
                            </button>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    );
}

export default Setup;
