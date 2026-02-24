const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
    const loginBtn = document.getElementById('login-btn');
    const uInput = document.getElementById('u-input');
    const pInput = document.getElementById('p-input');

    const errorMsg = document.getElementById('login-error');

    const attemptLogin = async () => {
        const u = uInput.value.trim();
        const p = pInput.value.trim();

        if (!u || !p) {
            await customAlert("Please enter both username and password.");
            return;
        }

        // Reset error
        errorMsg.style.display = 'none';

        try {
            const result = await ipcRenderer.invoke('login', { username: u, password: p });
            if (result.success) {
                // alert("Login successful! Redirecting..."); // Optional debug
                ipcRenderer.send('login-success', result.role);
            } else {
                errorMsg.textContent = "Invalid Username or Password";
                errorMsg.style.display = 'block';
                await customAlert("Login Failed: Invalid Username or Password"); // Explicit feedback

                // Shake animation for feedback
                const container = document.querySelector('.form-wrapper');
                container.style.transform = 'translateX(10px)';
                setTimeout(() => container.style.transform = 'translateX(-10px)', 100);
                setTimeout(() => container.style.transform = 'translateX(0)', 200);
            }
        } catch (err) {
            console.error(err);
            errorMsg.textContent = "An error occurred during login.";
            errorMsg.style.display = 'block';
            await customAlert("System Error: " + err.message); // Explicit feedback
        }
    };

    loginBtn.addEventListener('click', attemptLogin);

    // Allow Enter key to login
    pInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') attemptLogin();
    });

    // Password Toggle
    const togglePassword = document.getElementById('toggle-password');
    togglePassword.addEventListener('click', () => {
        const type = pInput.getAttribute('type') === 'password' ? 'text' : 'password';
        pInput.setAttribute('type', type);

        // Update Icon
        if (type === 'text') {
            togglePassword.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path><line x1="1" y1="1" x2="23" y2="23"></line></svg>`;
            togglePassword.style.color = '#3b82f6'; // Active color
        } else {
            togglePassword.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;
            togglePassword.style.color = '#cbd5e1'; // Default color
        }
    });

    initParticles();
});

// PARTICLE SYSTEM
function initParticles() {
    const canvas = document.getElementById('particle-canvas');
    const ctx = canvas.getContext('2d');
    const container = document.getElementById('brand-side');

    let width, height;
    let particles = [];

    function resize() {
        width = container.clientWidth;
        height = container.clientHeight;
        canvas.width = width;
        canvas.height = height;
    }

    window.addEventListener('resize', resize);
    resize();

    class Particle {
        constructor(x, y, type = 'normal') {
            this.x = x || Math.random() * width;
            this.y = y || Math.random() * height;
            this.vx = (Math.random() - 0.5) * 2;
            this.vy = (Math.random() - 0.5) * 2;
            this.size = Math.random() * 3 + 2; // Small balls
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

            // Bounce off walls
            if (this.x < 0 || this.x > width) this.vx *= -1;
            if (this.y < 0 || this.y > height) this.vy *= -1;

            // Burst particle lifecycle
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

    // Initial population
    for (let i = 0; i < 30; i++) {
        particles.push(new Particle());
    }

    // Animation Loop
    function animate() {
        ctx.clearRect(0, 0, width, height);

        particles.forEach((p, index) => {
            p.update();
            p.draw();
            if (p.type === 'burst' && p.life <= 0) {
                particles.splice(index, 1);
            }
        });

        requestAnimationFrame(animate);
    }
    animate();

    // Click Interaction (on container to catch clicks on text/logo)
    container.addEventListener('mousedown', (e) => {
        const rect = container.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;

        // Spawn burst
        for (let i = 0; i < 10; i++) {
            particles.push(new Particle(x, y, 'burst'));
        }
    });
}

// Custom alert replacement
function customAlert(message) {
    return new Promise((resolve) => {
        const modal = document.getElementById('custom-alert-modal');
        const messageEl = document.getElementById('custom-alert-message');
        const okBtn = document.getElementById('custom-alert-ok');

        messageEl.textContent = message;
        modal.style.display = 'flex';

        // Focus OK button for keyboard accessibility
        setTimeout(() => okBtn.focus(), 50);

        const cleanup = () => {
            modal.style.display = 'none';
            okBtn.removeEventListener('click', handleOk);
            document.removeEventListener('keydown', handleKeydown);
            resolve();
        };

        const handleOk = () => cleanup();
        const handleKeydown = (e) => {
            if (e.key === 'Enter' || e.key === 'Escape') handleOk();
        };

        okBtn.addEventListener('click', handleOk);
        document.addEventListener('keydown', handleKeydown);
    });
}
