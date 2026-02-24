import React, { useState, useEffect } from 'react';
import db from '../services/db';

function UserManagement({ showAlert }) {
    const [users, setUsers] = useState([]);
    const [newUser, setNewUser] = useState({
        username: '',
        fullName: '',
        email: '',
        role: 'teacher',
        password: ''
    });

    useEffect(() => {
        loadUsers();
    }, []);

    const loadUsers = async () => {
        const data = await db.getUsers();
        setUsers(data);
    };

    const generatePassword = () => {
        const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%';
        let password = '';
        for (let i = 0; i < 12; i++) {
            password += chars.charAt(Math.floor(Math.random() * chars.length));
        }
        setNewUser({ ...newUser, password });
    };

    const createUser = async () => {
        if (!newUser.username || !newUser.password) {
            await showAlert('Please fill in username and password.');
            return;
        }

        const userToCreate = {
            ...newUser,
            password: newUser.password,
            createdAt: new Date().toISOString()
        };

        const result = await db.createUser(userToCreate);
        if (result.success) {
            await showAlert('User created successfully!');
            setNewUser({ username: '', fullName: '', email: '', role: 'teacher', password: '' });
            loadUsers();
        } else {
            await showAlert(result.error || 'Failed to create user.');
        }
    };

    const deleteUser = async (username) => {
        if (username === 'admin') {
            await showAlert('Cannot delete the admin user.');
            return;
        }

        const result = await db.deleteUser(username);
        if (result.success) {
            loadUsers();
        }
    };

    return (
        <div id="user-management-section" className="user-management" style={{ display: 'block' }}>
            <h2>User Management</h2>

            <div className="create-user-form">
                <h3>Create New User</h3>
                <form onSubmit={(e) => { e.preventDefault(); createUser(); }}>
                    <div className="form-row">
                        <div className="form-group">
                            <label htmlFor="new-username">Username *</label>
                            <input
                                type="text"
                                id="new-username"
                                value={newUser.username}
                                onChange={(e) => setNewUser({ ...newUser, username: e.target.value })}
                                required
                            />
                        </div>
                        <div className="form-group">
                            <label htmlFor="new-fullname">Full Name</label>
                            <input
                                type="text"
                                id="new-fullname"
                                value={newUser.fullName}
                                onChange={(e) => setNewUser({ ...newUser, fullName: e.target.value })}
                            />
                        </div>
                    </div>
                    <div className="form-row">
                        <div className="form-group">
                            <label htmlFor="new-email">Email</label>
                            <input
                                type="email"
                                id="new-email"
                                value={newUser.email}
                                onChange={(e) => setNewUser({ ...newUser, email: e.target.value })}
                            />
                        </div>
                        <div className="form-group">
                            <label htmlFor="new-role">Role *</label>
                            <select
                                id="new-role"
                                value={newUser.role}
                                onChange={(e) => setNewUser({ ...newUser, role: e.target.value })}
                                required
                            >
                                <option value="teacher">Teacher</option>
                                <option value="admin">Admin</option>
                            </select>
                        </div>
                    </div>
                    <div className="form-row">
                        <div className="form-group">
                            <label htmlFor="new-password">Password *</label>
                            <input
                                type="text"
                                id="new-password"
                                value={newUser.password}
                                onChange={(e) => setNewUser({ ...newUser, password: e.target.value })}
                                required
                            />
                            <button type="button" className="generate-pwd-btn" onClick={generatePassword}>
                                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px' }}><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3y-3.5 3.5"></path></svg>
                                Generate Password
                            </button>
                        </div>
                    </div>
                    <button type="submit" className="btn btn-primary">Create User</button>
                </form>
            </div>

            <h3>All Users</h3>
            <div className="user-table">
                <table>
                    <thead>
                        <tr>
                            <th>Username</th>
                            <th>Full Name</th>
                            <th>Email</th>
                            <th>Role</th>
                            <th>Created</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {users.map((user) => (
                            <tr key={user.username}>
                                <td>{user.username}</td>
                                <td>{user.fullName || '-'}</td>
                                <td>{user.email || '-'}</td>
                                <td>{user.role}</td>
                                <td>{user.createdAt ? new Date(user.createdAt).toLocaleDateString() : '-'}</td>
                                <td>
                                    {user.username !== 'admin' && (
                                        <button
                                            className="btn-danger"
                                            onClick={() => deleteUser(user.username)}
                                            style={{ padding: '5px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '5px' }}
                                        >
                                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
                                            Delete
                                        </button>
                                    )}
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </div>
    );
}

export default UserManagement;
