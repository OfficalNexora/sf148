/**
 * db.js — Unified data service for Form 137 React App.
 * 
 * Detects runtime environment and routes calls to either:
 *   - Electron IPC (when running inside Electron)
 *   - IndexedDB + localStorage (when running in a normal browser)
 * 
 * All public methods mirror the Electron IPC channel signatures exactly,
 * making component migration a simple import swap.
 */

// --- Runtime Detection ---
let isElectron = false;
let ipcRenderer = null;
let nodeCrypto = null;

try {
    if (window && window.require) {
        const electron = window.require('electron');
        ipcRenderer = electron.ipcRenderer;
        nodeCrypto = window.require('crypto');
        isElectron = true;
    }
} catch (e) {
    // Not in Electron — browser mode
    isElectron = false;
}

// =====================================================================
// SHA-256 hashing (browser-compatible with fallback)
// =====================================================================

async function hashPassword(password) {
    // Electron: use Node crypto
    if (isElectron && nodeCrypto) {
        return nodeCrypto.createHash('sha256').update(password).digest('hex');
    }

    // Browser: prefer Web Crypto API
    if (typeof crypto !== 'undefined' && crypto.subtle) {
        try {
            const msgBuffer = new TextEncoder().encode(password);
            const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
            return Array.from(new Uint8Array(hashBuffer))
                .map(b => b.toString(16).padStart(2, '0')).join('');
        } catch (e) { /* fall through */ }
    }

    // Fallback: pure JS SHA-256
    return sha256Fallback(password);
}

function sha256Fallback(str) {
    function rightRotate(value, amount) {
        return (value >>> amount) | (value << (32 - amount));
    }
    const mathPow = Math.pow;
    const maxWord = mathPow(2, 32);
    let result = '';
    const words = [];
    const asciiBitLength = str.length * 8;
    let hash = [];
    const k = [];
    let primeCounter = 0;
    const isComposite = {};
    for (let candidate = 2; primeCounter < 64; candidate++) {
        if (!isComposite[candidate]) {
            for (let i = 0; i < 313; i += candidate) isComposite[i] = candidate;
            hash[primeCounter] = (mathPow(candidate, 0.5) * maxWord) | 0;
            k[primeCounter++] = (mathPow(candidate, 1 / 3) * maxWord) | 0;
        }
    }
    str += '\x80';
    while ((str.length % 64) - 56) str += '\x00';
    for (let i = 0; i < str.length; i++) {
        const j = str.charCodeAt(i);
        if (j >> 8) return '';
        words[i >> 2] |= j << (((3 - i) % 4) * 8);
    }
    words[words.length] = ((asciiBitLength / maxWord) | 0);
    words[words.length] = (asciiBitLength);
    for (let j = 0; j < words.length;) {
        const w = words.slice(j, j += 16);
        const oldHash = hash.slice(0);
        for (let i = 0; i < 64; i++) {
            const w15 = w[i - 15], w2 = w[i - 2];
            const a = hash[0], e = hash[4];
            const temp1 = hash[7]
                + (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25))
                + ((e & hash[5]) ^ ((~e) & hash[6]))
                + k[i]
                + (w[i] = (i < 16) ? w[i] : (
                    w[i - 16]
                    + (rightRotate(w15, 7) ^ rightRotate(w15, 18) ^ (w15 >>> 3))
                    + w[i - 7]
                    + (rightRotate(w2, 17) ^ rightRotate(w2, 19) ^ (w2 >>> 10))
                ) | 0);
            const temp2 = (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22))
                + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
            hash = [(temp1 + temp2) | 0].concat(hash);
            hash[4] = (hash[4] + temp1) | 0;
        }
        for (let i = 0; i < 8; i++) hash[i] = (hash[i] + oldHash[i]) | 0;
    }
    for (let i = 0; i < 8; i++) {
        for (let j = 3; j + 1; j--) {
            const b = (hash[i] >> (j * 8)) & 255;
            result += ((b < 16) ? '0' : '') + b.toString(16);
        }
    }
    return result;
}

// =====================================================================
// Deep Merge Helper
// =====================================================================
function deepMerge(target, source) {
    if (typeof target !== 'object' || target === null) return source;
    if (typeof source !== 'object' || source === null) return target;

    const output = { ...target };
    Object.keys(source).forEach(key => {
        if (Array.isArray(source[key])) {
            const existingArray = output[key] || [];
            // Merge arrays by ID (students)
            const idMap = new Map();
            existingArray.forEach(item => idMap.set(item.id, item));
            source[key].forEach(item => idMap.set(item.id, item));
            output[key] = Array.from(idMap.values());
        } else if (typeof source[key] === 'object' && source[key] !== null) {
            output[key] = deepMerge(target[key] || {}, source[key]);
        } else {
            output[key] = source[key];
        }
    });
    return output;
}

// =====================================================================
// IndexedDB Backend (browser mode only)
// =====================================================================

class BrowserDB {
    constructor() {
        this.dbName = 'Form137App';
        this.dbVersion = 1;
        this.db = null;
        this.ready = this._init();
    }

    async _init() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(this.dbName, this.dbVersion);

            request.onerror = (e) => {
                console.error('IndexedDB error:', e.target.error);
                reject(e.target.error);
            };

            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains('users'))
                    db.createObjectStore('users', { keyPath: 'username' });
                if (!db.objectStoreNames.contains('structure'))
                    db.createObjectStore('structure', { keyPath: 'id' });
                if (!db.objectStoreNames.contains('records'))
                    db.createObjectStore('records', { keyPath: 'id' });
            };

            request.onsuccess = async (e) => {
                this.db = e.target.result;
                await this._ensureDefaults();
                resolve();
            };
        });
    }

    async _ensureDefaults() {
        // Seed default structure if it doesn't exist
        const structure = await this._get('structure', 'main');
        if (!structure) {
            await this._put('structure', {
                id: 'main',
                data: { "Grade 11": { "TVL - ICT": { "Section A": [] } } }
            });
        }
    }

    _get(store, key) {
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction([store], 'readonly');
            const req = tx.objectStore(store).get(key);
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    _getAll(store) {
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction([store], 'readonly');
            const req = tx.objectStore(store).getAll();
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    _put(store, value) {
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction([store], 'readwrite');
            const req = tx.objectStore(store).put(value);
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    _delete(store, key) {
        return new Promise((resolve, reject) => {
            const tx = this.db.transaction([store], 'readwrite');
            const req = tx.objectStore(store).delete(key);
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }
}

// Singleton — only created in browser mode
let browserDB = null;
function getBrowserDB() {
    if (!browserDB) browserDB = new BrowserDB();
    return browserDB;
}

// =====================================================================
// Unified Public API
// =====================================================================

const db = {
    // --- Auth & Setup ---
    async checkNeedsSetup() {
        if (isElectron) {
            return ipcRenderer.invoke('check-needs-setup');
        }
        const bdb = getBrowserDB();
        await bdb.ready;
        const users = await bdb._getAll('users');
        return users.length === 0;
    },

    async setupAdmin(username, password) {
        if (isElectron) {
            const result = await ipcRenderer.invoke('setup-admin', { username, password });
            if (result.success) {
                localStorage.setItem('currentUserRole', result.role);
            }
            return result;
        }

        const bdb = getBrowserDB();
        await bdb.ready;
        const users = await bdb._getAll('users');
        if (users.length > 0) {
            return { success: false, error: 'Setup already complete.' };
        }

        const adminUser = {
            username,
            password: await hashPassword(password),
            role: 'admin',
            fullName: 'System Administrator',
            createdAt: new Date().toISOString()
        };

        await bdb._put('users', adminUser);
        localStorage.setItem('currentUserRole', 'admin');
        return { success: true, role: 'admin' };
    },

    async login(username, password) {
        if (isElectron) {
            const result = await ipcRenderer.invoke('login', { username, password });
            if (result.success) {
                localStorage.setItem('currentUserRole', result.role);
            }
            return result;
        }

        const bdb = getBrowserDB();
        await bdb.ready;
        const users = await bdb._getAll('users');
        for (const u of users) {
            if (u.username === username) {
                if (u.password.length === 64) {
                    const hashed = await hashPassword(password);
                    if (u.password === hashed) {
                        localStorage.setItem('currentUserRole', u.role);
                        return { success: true, role: u.role };
                    }
                } else {
                    if (u.password === password) {
                        localStorage.setItem('currentUserRole', u.role);
                        return { success: true, role: u.role };
                    }
                }
            }
        }
        return { success: false };
    },

    async logout() {
        localStorage.removeItem('currentUserRole');
        if (isElectron) {
            await ipcRenderer.invoke('logout');
        }
        return { success: true };
    },

    async getCurrentUserRole() {
        // Always check localStorage first (works in both modes)
        const stored = localStorage.getItem('currentUserRole');
        if (stored) return stored;

        if (isElectron) {
            const role = await ipcRenderer.invoke('get-current-user-role');
            if (role) localStorage.setItem('currentUserRole', role);
            return role;
        }
        return null;
    },

    // --- Structure ---
    async getStructure() {
        if (isElectron) {
            return ipcRenderer.invoke('get-structure');
        }
        const bdb = getBrowserDB();
        await bdb.ready;
        const record = await bdb._get('structure', 'main');
        return record ? record.data : null;
    },

    async updateStructure(newStructure) {
        if (isElectron) {
            return ipcRenderer.invoke('update-structure', newStructure);
        }
        const bdb = getBrowserDB();
        await bdb.ready;
        await bdb._put('structure', { id: 'main', data: newStructure });
        return { success: true };
    },

    // --- Student Records ---
    async loadStudent(id) {
        if (isElectron) {
            return ipcRenderer.invoke('load-student', id);
        }
        const bdb = getBrowserDB();
        await bdb.ready;
        const record = await bdb._get('records', id);
        return record ? record.data : null;
    },

    async saveStudent(id, data) {
        if (isElectron) {
            return ipcRenderer.invoke('save-student', { id, data });
        }
        const bdb = getBrowserDB();
        await bdb.ready;
        try {
            await bdb._put('records', { id, data });
            return { success: true };
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    // --- User Management ---
    async getUsers() {
        if (isElectron) {
            return ipcRenderer.invoke('get-users');
        }
        const bdb = getBrowserDB();
        await bdb.ready;
        const users = await bdb._getAll('users');
        return users.map(({ password, ...safe }) => safe);
    },

    async createUser(newUser) {
        if (isElectron) {
            return ipcRenderer.invoke('create-user', newUser);
        }
        const bdb = getBrowserDB();
        await bdb.ready;

        if (newUser.role !== 'teacher' && newUser.role !== 'admin') {
            return { success: false, error: 'Invalid role.' };
        }

        const users = await bdb._getAll('users');
        if (users.find(u => u.username === newUser.username)) {
            return { success: false, error: 'User already exists.' };
        }

        const { plainPassword, ...userToSave } = newUser;
        if (userToSave.password) {
            userToSave.password = await hashPassword(userToSave.password);
        }
        await bdb._put('users', userToSave);
        const { password, ...safeUser } = userToSave;
        return { success: true, user: safeUser };
    },

    async deleteUser(username) {
        if (isElectron) {
            return ipcRenderer.invoke('delete-user', username);
        }
        const bdb = getBrowserDB();
        await bdb.ready;
        await bdb._delete('users', username);
        return { success: true };
    },

    // --- File Operations (browser fallbacks) ---
    async importStudentFromFile() {
        // Browser mode: use file input
        return new Promise((resolve) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.json';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) { resolve(null); return; }
                try {
                    const text = await file.text();
                    resolve(JSON.parse(text));
                } catch (err) {
                    resolve(null);
                }
            };
            input.click();
        });
    },

    async exportStudentToFile(data, filename) {
        // Browser mode: download as file
        const json = JSON.stringify(data, null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
        return { success: true };
    },

    // --- DB Syncing (Hub & Spoke) ---
    async exportSyncFile() {
        if (isElectron) {
            return ipcRenderer.invoke('export-sync');
        }
        const bdb = getBrowserDB();
        await bdb.ready;

        const structureRecord = await bdb._get('structure', 'main');
        const structure = structureRecord ? structureRecord.data : {};
        const records = await bdb._getAll('records');

        const syncData = {
            version: 1,
            timestamp: new Date().toISOString(),
            structure,
            records: records.map(r => r.data)
        };

        const json = JSON.stringify(syncData, null, 2);
        const blob = new Blob([json], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const dateStr = new Date().toISOString().split('T')[0];
        a.download = `form137_sync_${dateStr}.f137sync`;
        a.click();
        URL.revokeObjectURL(url);

        return { success: true };
    },

    async importSyncFile() {
        if (isElectron) {
            return ipcRenderer.invoke('import-sync');
        }

        // Browser mode: open file picker
        return new Promise((resolve) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.f137sync,.json';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) { resolve({ success: false, error: 'No file selected' }); return; }
                try {
                    const text = await file.text();
                    const syncData = JSON.parse(text);

                    if (!syncData.structure || !syncData.records) {
                        resolve({ success: false, error: 'Invalid sync file format' });
                        return;
                    }

                    const bdb = getBrowserDB();
                    await bdb.ready;

                    // 1. Merge Structure
                    const currentStructRec = await bdb._get('structure', 'main');
                    const currentStruct = currentStructRec ? currentStructRec.data : {};

                    // Simple merge: overwrite current with incoming (for sync purpose)
                    // Or ideally deep merge, but for now we'll do an overwrite of existing branches, but keep local refs?
                    // Actually, the user approved "incoming overwrites".
                    // But an absolute overwrite of the WHOLE structure might wipe out local Admin additions.
                    // Let's do a basic deep merge.
                    const mergedStructure = deepMerge(currentStruct, syncData.structure);
                    await bdb._put('structure', { id: 'main', data: mergedStructure });

                    // 2. Merge Records
                    for (const recordData of syncData.records) {
                        const id = recordData.id || recordData.info?.lrn;
                        if (id) {
                            await bdb._put('records', { id, data: recordData });
                        }
                    }

                    resolve({ success: true, count: syncData.records.length });
                } catch (err) {
                    resolve({ success: false, error: err.message });
                }
            };
            input.click();
        });
    },

    /**
     * Merge a sync data object directly (used by online SyncInbox after decoding from Firestore).
     * @param {object} syncData - { structure, records[] }
     */
    async importSyncData(syncData) {
        if (!syncData || !syncData.structure || !syncData.records) {
            return { success: false, error: 'Invalid sync data format' };
        }

        if (isElectron) {
            // Electron: write structure + records directly to disk
            try {
                const fs = window.require('fs');
                const path = window.require('path');
                const DATA_DIR = path.join(__dirname, '..', '..', 'data');
                const RECORDS_DIR = path.join(DATA_DIR, 'records');

                const structPath = path.join(DATA_DIR, 'structure.json');
                let currentStruct = {};
                if (fs.existsSync(structPath)) {
                    currentStruct = JSON.parse(fs.readFileSync(structPath));
                }

                const mergedStructure = deepMerge(currentStruct, syncData.structure);
                fs.writeFileSync(structPath, JSON.stringify(mergedStructure, null, 2));

                for (const recordData of syncData.records) {
                    const id = recordData.id || recordData.info?.lrn;
                    if (id) {
                        const recPath = path.join(RECORDS_DIR, `${id}.json`);
                        let mergedRecord = recordData;
                        if (fs.existsSync(recPath)) {
                            const existing = JSON.parse(fs.readFileSync(recPath));
                            mergedRecord = deepMerge(existing, recordData);
                        }
                        fs.writeFileSync(recPath, JSON.stringify(mergedRecord, null, 2));
                    }
                }
                return { success: true, count: syncData.records.length };
            } catch (err) {
                console.error('Electron importSyncData error:', err);
                return { success: false, error: err.message };
            }
        }

        // Browser mode: IndexedDB
        const bdb = getBrowserDB();
        await bdb.ready;

        const currentStructRec = await bdb._get('structure', 'main');
        const currentStruct = currentStructRec ? currentStructRec.data : {};
        const mergedStructure = deepMerge(currentStruct, syncData.structure);
        await bdb._put('structure', { id: 'main', data: mergedStructure });

        for (const recordData of syncData.records) {
            const id = recordData.id || recordData.info?.lrn;
            if (id) {
                const existingRec = await bdb._get('records', id);
                let mergedRecord = recordData;
                if (existingRec && existingRec.data) {
                    mergedRecord = deepMerge(existingRec.data, recordData);
                }
                await bdb._put('records', { id, data: mergedRecord });
            }
        }
        return { success: true, count: syncData.records.length };
    },

    // --- Utility ---
    isElectron() {
        return isElectron;
    },

    // Expose the BrowserDB instance for direct access (e.g. gather all records for sync)
    _getBrowserDB() {
        if (isElectron) return null;
        return getBrowserDB();
    },

    hashPassword
};

export default db;
