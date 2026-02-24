// User Management Frontend Logic
// ipcRenderer is already declared in dashboard.js

// Elements
const userTableBody = document.getElementById('user-table-body');
const createUserForm = document.getElementById('create-user-form');
const generatePwdBtn = document.getElementById('generate-pwd-btn');
const createUserBtn = document.getElementById('create-user-btn');

// Load users on init
function loadUsers() {
    ipcRenderer.send('get-users');
}

ipcRenderer.on('users-loaded', (event, users) => {
    renderUserTable(users);
});

function renderUserTable(users) {
    if (!userTableBody) return;
    userTableBody.innerHTML = '';
    users.forEach(user => {
        const row = document.createElement('tr');

        const usernameCell = document.createElement('td');
        usernameCell.textContent = user.username;
        row.appendChild(usernameCell);

        const fullNameCell = document.createElement('td');
        fullNameCell.textContent = user.fullName || '';
        row.appendChild(fullNameCell);

        const emailCell = document.createElement('td');
        emailCell.textContent = user.email || '';
        row.appendChild(emailCell);

        const roleCell = document.createElement('td');
        roleCell.textContent = user.role;
        row.appendChild(roleCell);

        const createdCell = document.createElement('td');
        createdCell.textContent = user.createdAt ? new Date(user.createdAt).toLocaleDateString() : '';
        row.appendChild(createdCell);

        const actionsCell = document.createElement('td');
        actionsCell.innerHTML = `
            <button onclick="deleteUser('${user.username}')">
                <i class="fas fa-trash-alt"></i> Delete
            </button>
        `;
        row.appendChild(actionsCell);
        userTableBody.appendChild(row);
    });
}

// Delete user function
function deleteUser(username) {
    if (confirm(`Delete user "${username}"?`)) {
        ipcRenderer.send('delete-user', username);
    }
}

// Helper to escape HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Generate password
function generatePassword() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!@#$%';
    let pwd = '';
    for (let i = 0; i < 12; i++) {
        pwd += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return pwd;
}

if (generatePwdBtn) {
    generatePwdBtn.addEventListener('click', () => {
        const pwd = generatePassword();
        const pwdInput = document.getElementById('new-password');
        if (pwdInput) pwdInput.value = pwd;
    });
}

if (createUserBtn) {
    createUserBtn.addEventListener('click', async () => {
        const username = document.getElementById('new-username').value.trim();
        const fullName = document.getElementById('new-fullname').value.trim();
        const email = document.getElementById('new-email').value.trim();
        const role = document.getElementById('new-role').value;
        const password = document.getElementById('new-password').value;
        if (!username || !password) {
            await customAlert('Username and password are required');
            return;
        }
        // Role validation – only teacher or admin allowed
        if (role !== 'teacher' && role !== 'admin') {
            await customAlert('Invalid role selected. Only Teacher or Admin are allowed.');
            return;
        }
        // Build user object (plainPassword is temporary for printing)
        const newUser = {
            username,
            fullName,
            email,
            role,
            password: hashPassword(password), // store hashed password
            plainPassword: password, // keep for immediate printing if needed
            createdAt: new Date().toISOString()
        };
        ipcRenderer.send('create-user', newUser);
        // Reset form
        createUserForm.reset();
    });
}

// Listen for creation/deletion updates
ipcRenderer.on('user-created', (event, user) => {
    loadUsers();
});
ipcRenderer.on('user-deleted', (event, username) => {
    loadUsers();
});

// Password hashing (same as backend) – duplicated for consistency
function hashPassword(password) {
    // Simple SHA-256 via Node crypto (available in Electron renderer)
    const crypto = require('crypto');
    return crypto.createHash('sha256').update(password).digest('hex');
}

// Initialize
loadUsers();
