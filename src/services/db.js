/**
 * Unified data service for Form 137 React App.
 *
 * Runtime routing:
 * - Electron: IPC channels handled by electron/main.js
 * - Browser: Firestore-backed shared data (users, structure, records)
 */

import {
    collection,
    deleteDoc,
    doc,
    getDoc,
    getDocs,
    limit,
    query,
    setDoc
} from 'firebase/firestore';
import { getFirebaseDb } from './firebaseClient';

// Runtime detection
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
    isElectron = false;
}

const STORAGE_KEYS = {
    role: 'currentUserRole',
    username: 'currentUsername'
};

const DEFAULT_STRUCTURE = {
    'Grade 11': {
        'TVL - ICT': {
            'Section A': []
        }
    }
};

const USERS_COLLECTION = 'users';
const RECORDS_COLLECTION = 'records';
const APPDATA_COLLECTION = 'appData';
const STRUCTURE_DOC_ID = 'structure';

function getBrowserFirestore() {
    return getFirebaseDb();
}

async function hashPassword(password) {
    if (isElectron && nodeCrypto) {
        return nodeCrypto.createHash('sha256').update(password).digest('hex');
    }

    if (typeof crypto !== 'undefined' && crypto.subtle) {
        try {
            const msgBuffer = new TextEncoder().encode(password);
            const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
            return Array.from(new Uint8Array(hashBuffer))
                .map(b => b.toString(16).padStart(2, '0')).join('');
        } catch (e) {
            // Fall through to JS fallback.
        }
    }

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
            const w15 = w[i - 15];
            const w2 = w[i - 2];
            const a = hash[0];
            const e = hash[4];
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

function deepMerge(target, source) {
    if (typeof target !== 'object' || target === null) return source;
    if (typeof source !== 'object' || source === null) return target;

    const output = { ...target };
    Object.keys(source).forEach(key => {
        if (Array.isArray(source[key])) {
            const existingArray = output[key] || [];
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

async function ensureStructureDoc(firestore) {
    const structureRef = doc(firestore, APPDATA_COLLECTION, STRUCTURE_DOC_ID);
    const structureSnap = await getDoc(structureRef);

    if (structureSnap.exists()) {
        return structureSnap.data().data || null;
    }

    await setDoc(structureRef, { data: DEFAULT_STRUCTURE });
    return DEFAULT_STRUCTURE;
}

const db = {
    async checkNeedsSetup() {
        if (isElectron) {
            return ipcRenderer.invoke('check-needs-setup');
        }

        const firestore = getBrowserFirestore();
        const usersRef = collection(firestore, USERS_COLLECTION);
        const usersSnap = await getDocs(query(usersRef, limit(1)));
        return usersSnap.empty;
    },

    async setupAdmin(username, password) {
        if (isElectron) {
            const result = await ipcRenderer.invoke('setup-admin', { username, password });
            if (result.success) {
                localStorage.setItem(STORAGE_KEYS.role, result.role);
                localStorage.setItem(STORAGE_KEYS.username, username);
            }
            return result;
        }

        try {
            const firestore = getBrowserFirestore();
            const usersRef = collection(firestore, USERS_COLLECTION);
            const usersSnap = await getDocs(query(usersRef, limit(1)));
            if (!usersSnap.empty) {
                return { success: false, error: 'Setup already complete.' };
            }

            const adminUser = {
                username,
                password: await hashPassword(password),
                role: 'admin',
                fullName: 'System Administrator',
                createdAt: new Date().toISOString()
            };

            await setDoc(doc(firestore, USERS_COLLECTION, username), adminUser);
            localStorage.setItem(STORAGE_KEYS.role, 'admin');
            localStorage.setItem(STORAGE_KEYS.username, username);
            return { success: true, role: 'admin' };
        } catch (err) {
            return { success: false, error: err.message };
        }
    },

    async login(username, password) {
        if (isElectron) {
            const result = await ipcRenderer.invoke('login', { username, password });
            if (result.success) {
                localStorage.setItem(STORAGE_KEYS.role, result.role);
                localStorage.setItem(STORAGE_KEYS.username, username);
            }
            return result;
        }

        try {
            const firestore = getBrowserFirestore();
            const userRef = doc(firestore, USERS_COLLECTION, username);
            const userSnap = await getDoc(userRef);
            if (!userSnap.exists()) {
                return { success: false };
            }

            const user = userSnap.data();
            let isValid = false;
            if (user.password && user.password.length === 64) {
                const hashed = await hashPassword(password);
                isValid = user.password === hashed;
            } else {
                isValid = user.password === password;
            }

            if (!isValid) {
                return { success: false };
            }

            localStorage.setItem(STORAGE_KEYS.role, user.role);
            localStorage.setItem(STORAGE_KEYS.username, user.username || username);
            return { success: true, role: user.role };
        } catch (err) {
            return { success: false, error: err.message };
        }
    },

    async logout() {
        localStorage.removeItem(STORAGE_KEYS.role);
        localStorage.removeItem(STORAGE_KEYS.username);

        if (isElectron) {
            await ipcRenderer.invoke('logout');
        }

        return { success: true };
    },

    async getCurrentUserRole() {
        const storedRole = localStorage.getItem(STORAGE_KEYS.role);
        const storedUsername = localStorage.getItem(STORAGE_KEYS.username);

        if (isElectron) {
            if (storedRole) return storedRole;
            const role = await ipcRenderer.invoke('get-current-user-role');
            if (role) localStorage.setItem(STORAGE_KEYS.role, role);
            return role;
        }

        if (!storedRole || !storedUsername) return null;

        try {
            const firestore = getBrowserFirestore();
            const userSnap = await getDoc(doc(firestore, USERS_COLLECTION, storedUsername));
            if (!userSnap.exists()) {
                localStorage.removeItem(STORAGE_KEYS.role);
                localStorage.removeItem(STORAGE_KEYS.username);
                return null;
            }

            const user = userSnap.data();
            if (user.role !== storedRole) {
                localStorage.setItem(STORAGE_KEYS.role, user.role || 'teacher');
                return user.role || 'teacher';
            }

            return storedRole;
        } catch {
            return null;
        }
    },

    async getStructure() {
        if (isElectron) {
            return ipcRenderer.invoke('get-structure');
        }

        try {
            const firestore = getBrowserFirestore();
            return await ensureStructureDoc(firestore);
        } catch (err) {
            console.error('getStructure failed:', err);
            return DEFAULT_STRUCTURE;
        }
    },

    async updateStructure(newStructure) {
        if (isElectron) {
            return ipcRenderer.invoke('update-structure', newStructure);
        }

        try {
            const firestore = getBrowserFirestore();
            await setDoc(doc(firestore, APPDATA_COLLECTION, STRUCTURE_DOC_ID), { data: newStructure });
            return { success: true };
        } catch (err) {
            return { success: false, error: err.message };
        }
    },

    async loadStudent(id) {
        if (isElectron) {
            return ipcRenderer.invoke('load-student', id);
        }

        const firestore = getBrowserFirestore();
        const recordSnap = await getDoc(doc(firestore, RECORDS_COLLECTION, id));
        if (!recordSnap.exists()) return null;

        const data = recordSnap.data();
        if (!data.id) return { ...data, id };
        return data;
    },

    async saveStudent(id, data) {
        if (isElectron) {
            return ipcRenderer.invoke('save-student', { id, data });
        }

        try {
            const firestore = getBrowserFirestore();
            const payload = { ...(data || {}) };
            if (!payload.id) payload.id = id;
            await setDoc(doc(firestore, RECORDS_COLLECTION, id), payload);
            return { success: true };
        } catch (error) {
            return { success: false, error: error.message };
        }
    },

    async getAllRecords() {
        if (isElectron) {
            return ipcRenderer.invoke('get-all-records');
        }

        const firestore = getBrowserFirestore();
        const snap = await getDocs(collection(firestore, RECORDS_COLLECTION));
        return snap.docs.map(d => {
            const item = d.data();
            if (!item.id) return { ...item, id: d.id };
            return item;
        });
    },

    async getUsers() {
        if (isElectron) {
            return ipcRenderer.invoke('get-users');
        }

        const firestore = getBrowserFirestore();
        const snap = await getDocs(collection(firestore, USERS_COLLECTION));
        return snap.docs.map(d => {
            const user = d.data();
            const { password, plainPassword, ...safe } = user;
            return safe;
        });
    },

    async createUser(newUser) {
        if (isElectron) {
            return ipcRenderer.invoke('create-user', newUser);
        }

        try {
            if (newUser.role !== 'teacher' && newUser.role !== 'admin') {
                return { success: false, error: 'Invalid role.' };
            }

            const firestore = getBrowserFirestore();
            const username = (newUser.username || '').trim();
            if (!username) {
                return { success: false, error: 'Username is required.' };
            }

            const userRef = doc(firestore, USERS_COLLECTION, username);
            const existing = await getDoc(userRef);
            if (existing.exists()) {
                return { success: false, error: 'User already exists.' };
            }

            const { plainPassword, ...userToSave } = newUser;
            userToSave.username = username;
            userToSave.createdAt = userToSave.createdAt || new Date().toISOString();

            if (userToSave.password) {
                if (userToSave.password.length !== 64) {
                    userToSave.password = await hashPassword(userToSave.password);
                }
            } else {
                return { success: false, error: 'Password is required.' };
            }

            await setDoc(userRef, userToSave);
            const { password, ...safeUser } = userToSave;
            return { success: true, user: safeUser };
        } catch (err) {
            return { success: false, error: err.message };
        }
    },

    async deleteUser(username) {
        if (isElectron) {
            return ipcRenderer.invoke('delete-user', username);
        }

        try {
            const firestore = getBrowserFirestore();
            await deleteDoc(doc(firestore, USERS_COLLECTION, username));
            return { success: true };
        } catch (err) {
            return { success: false, error: err.message };
        }
    },

    async importStudentFromFile() {
        return new Promise((resolve) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.json';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) {
                    resolve(null);
                    return;
                }
                try {
                    const text = await file.text();
                    resolve(JSON.parse(text));
                } catch {
                    resolve(null);
                }
            };
            input.click();
        });
    },

    async exportStudentToFile(data, filename) {
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

    async exportSyncFile() {
        if (isElectron) {
            return ipcRenderer.invoke('export-sync');
        }

        const structure = await this.getStructure();
        const records = await this.getAllRecords();

        const syncData = {
            version: 1,
            timestamp: new Date().toISOString(),
            structure,
            records
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

        return new Promise((resolve) => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.f137sync,.json';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) {
                    resolve({ success: false, error: 'No file selected' });
                    return;
                }

                try {
                    const text = await file.text();
                    const syncData = JSON.parse(text);
                    const result = await this.importSyncData(syncData);
                    resolve(result);
                } catch (err) {
                    resolve({ success: false, error: err.message });
                }
            };
            input.click();
        });
    },

    async importSyncData(syncData) {
        if (!syncData || !syncData.structure || !syncData.records) {
            return { success: false, error: 'Invalid sync data format' };
        }

        if (isElectron) {
            return ipcRenderer.invoke('import-sync-data', syncData);
        }

        const currentStruct = await this.getStructure();
        const mergedStructure = deepMerge(currentStruct || {}, syncData.structure);
        const structResult = await this.updateStructure(mergedStructure);
        if (!structResult.success) {
            return structResult;
        }

        let savedCount = 0;
        for (const recordData of syncData.records) {
            const id = recordData.id || recordData.info?.lrn;
            if (!id) continue;

            const existingRecord = await this.loadStudent(id);
            const mergedRecord = existingRecord
                ? deepMerge(existingRecord, recordData)
                : recordData;

            const result = await this.saveStudent(id, mergedRecord);
            if (result.success) savedCount++;
        }

        return { success: true, count: savedCount };
    },

    isElectron() {
        return isElectron;
    },

    hashPassword
};

export default db;
