const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const xlsx = require('xlsx');

// DIRECTORY SETUP - Use parent folder's data directory for continuity
const DATA_DIR = path.join(__dirname, '..', '..', 'data');
const RECORDS_DIR = path.join(DATA_DIR, 'records');
const ASSETS_DIR = path.join(__dirname, '..', 'public', 'assets');

// Ensure directories exist
[DATA_DIR, RECORDS_DIR].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

let mainWindow;
let currentUserRole = null;

// Check if we're in development
const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1366,
        height: 768,
        icon: path.join(__dirname, '..', 'public', 'assets', 'images', 'deped_logo.webp'),
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            backgroundThrottling: false
        }
    });

    if (isDev) {
        // Development: load from Vite dev server
        mainWindow.loadURL('http://localhost:5173');
        mainWindow.webContents.openDevTools();
    } else {
        // Production: load built files
        mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
    }

    mainWindow.setMenuBarVisibility(false);

    mainWindow.webContents.on('did-create-window', (childWindow) => {
        childWindow.setMenuBarVisibility(false);
    });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

// --- PASSWORD HASHING ---
function hashPassword(password) {
    return crypto.createHash('sha256').update(password).digest('hex');
}

// --- DEEP MERGE HELPER ---
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

// --- IPC HANDLERS ---

ipcMain.handle('get-current-user-role', () => currentUserRole);

// CHECK NEEDS SETUP
ipcMain.handle('check-needs-setup', async () => {
    const usersPath = path.join(DATA_DIR, 'users.json');
    if (!fs.existsSync(usersPath)) return true;
    try {
        const users = JSON.parse(fs.readFileSync(usersPath));
        return users.length === 0;
    } catch {
        return true;
    }
});

// SETUP ADMIN
ipcMain.handle('setup-admin', async (event, { username, password }) => {
    const usersPath = path.join(DATA_DIR, 'users.json');

    // Safety check - only allow if empty
    let users = [];
    if (fs.existsSync(usersPath)) {
        try { users = JSON.parse(fs.readFileSync(usersPath)); } catch { }
    }

    if (users.length > 0) {
        return { success: false, error: 'Setup already complete.' };
    }

    const adminUser = {
        username: username,
        password: hashPassword(password),
        role: 'admin',
        fullName: 'System Administrator',
        createdAt: new Date().toISOString()
    };

    fs.writeFileSync(usersPath, JSON.stringify([adminUser], null, 2));
    currentUserRole = 'admin';
    return { success: true, role: 'admin' };
});

// LOGIN
ipcMain.handle('login', async (event, { username, password }) => {
    const usersPath = path.join(DATA_DIR, 'users.json');
    const usersJsPath = path.join(DATA_DIR, 'users.js');

    // Migration from users.js to users.json
    if (!fs.existsSync(usersPath)) {
        if (fs.existsSync(usersJsPath)) {
            try {
                const raw = fs.readFileSync(usersJsPath, 'utf8');
                const parsed = JSON.parse(raw);
                if (Array.isArray(parsed)) {
                    fs.writeFileSync(usersPath, JSON.stringify(parsed, null, 2));
                }
            } catch (e) {
                console.error('Failed to parse users.js fallback as JSON:', e);
            }
        } else {
            // Start empty
            fs.writeFileSync(usersPath, JSON.stringify([], null, 2));
        }
    }

    const users = JSON.parse(fs.readFileSync(usersPath));
    const validUser = users.find(u => {
        if (u.password.length === 64) {
            return u.username === username && u.password === hashPassword(password);
        } else {
            return u.username === username && u.password === password;
        }
    });

    if (validUser) {
        currentUserRole = validUser.role;
        return { success: true, role: validUser.role };
    }
    return { success: false };
});

// LOGOUT
ipcMain.handle('logout', () => {
    currentUserRole = null;
    return { success: true };
});

// GET TREE STRUCTURE
ipcMain.handle('get-structure', async () => {
    const structPath = path.join(DATA_DIR, 'structure.json');
    if (!fs.existsSync(structPath)) {
        const dummy = { "Grade 11": { "TVL - ICT": { "Section A": [] } } };
        fs.writeFileSync(structPath, JSON.stringify(dummy, null, 2));
        return dummy;
    }
    return JSON.parse(fs.readFileSync(structPath));
});

// UPDATE STRUCTURE
ipcMain.handle('update-structure', async (event, newStructure) => {
    const structPath = path.join(DATA_DIR, 'structure.json');
    fs.writeFileSync(structPath, JSON.stringify(newStructure, null, 2));
    return { success: true };
});

// LOAD STUDENT RECORD
ipcMain.handle('load-student', async (event, id) => {
    const recordPath = path.join(RECORDS_DIR, `${id}.json`);
    if (fs.existsSync(recordPath)) {
        return JSON.parse(fs.readFileSync(recordPath));
    }
    return null;
});

// SAVE STUDENT RECORD
ipcMain.handle('save-student', async (event, { id, data }) => {
    try {
        const recordPath = path.join(RECORDS_DIR, `${id}.json`);
        fs.writeFileSync(recordPath, JSON.stringify(data, null, 2));
        return { success: true };
    } catch (error) {
        console.error("Save failed:", error);
        return { success: false, error: error.message };
    }
});

// GET ALL STUDENT RECORDS
ipcMain.handle('get-all-records', async () => {
    try {
        if (!fs.existsSync(RECORDS_DIR)) return [];
        const files = fs.readdirSync(RECORDS_DIR).filter(f => f.endsWith('.json'));
        const records = [];
        for (const file of files) {
            try {
                const parsed = JSON.parse(fs.readFileSync(path.join(RECORDS_DIR, file)));
                records.push(parsed);
            } catch (err) {
                console.error('Failed reading record file:', file, err.message);
            }
        }
        return records;
    } catch (err) {
        console.error('Failed to get all records:', err);
        return [];
    }
});

// USB SCANNER (Windows)
ipcMain.handle('scan-usb', async () => {
    const drives = [];
    if (process.platform === 'win32') {
        const letters = "DEFGHIJKLMNOPQRSTUVWXYZ".split('');
        for (const letter of letters) {
            const drivePath = `${letter}:\\`;
            try {
                if (fs.existsSync(drivePath)) {
                    drives.push({ path: drivePath, label: `Drive (${letter}:)` });
                }
            } catch (e) { /* Drive not ready */ }
        }
    }
    return drives;
});

// EXPORT FILE
ipcMain.handle('export-file', async (event, { path: drivePath, filename, content }) => {
    try {
        const fullPath = path.join(drivePath, filename);
        fs.writeFileSync(fullPath, content);
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
});

// FILE DIALOGS
ipcMain.handle('save-file-dialog', async (event, defaultName) => {
    const { filePath } = await dialog.showSaveDialog({
        title: 'Export Student Record',
        defaultPath: defaultName,
        filters: [{ name: 'JSON Files', extensions: ['json'] }]
    });
    return filePath;
});

ipcMain.handle('open-file-dialog', async () => {
    const { filePaths } = await dialog.showOpenDialog({
        title: 'Import Student Record',
        filters: [{ name: 'JSON Files', extensions: ['json'] }],
        properties: ['openFile']
    });
    return filePaths[0];
});

ipcMain.handle('read-file', async (event, filePath) => {
    try {
        return fs.readFileSync(filePath, 'utf8');
    } catch (e) {
        return null;
    }
});

// USER MANAGEMENT
ipcMain.handle('get-users', async () => {
    const usersPath = path.join(DATA_DIR, 'users.json');
    if (!fs.existsSync(usersPath)) {
        const defaultUsers = [{
            username: 'admin',
            password: hashPassword('admin'),
            role: 'admin',
            fullName: 'System Administrator',
            createdAt: new Date().toISOString()
        }];
        fs.writeFileSync(usersPath, JSON.stringify(defaultUsers, null, 2));
    }
    const users = JSON.parse(fs.readFileSync(usersPath));
    // Don't send password hashes to frontend
    return users.map(u => {
        const { password, plainPassword, ...safe } = u;
        return safe;
    });
});

ipcMain.handle('create-user', async (event, newUser) => {
    if (newUser.role !== 'teacher' && newUser.role !== 'admin') {
        return { success: false, error: 'Invalid role. Only teacher or admin allowed.' };
    }

    const usersPath = path.join(DATA_DIR, 'users.json');
    const users = JSON.parse(fs.readFileSync(usersPath));

    const { plainPassword, ...userToSave } = newUser;
    users.push(userToSave);

    fs.writeFileSync(usersPath, JSON.stringify(users, null, 2));

    const { password, ...safeUser } = userToSave;
    return { success: true, user: safeUser };
});

ipcMain.handle('delete-user', async (event, username) => {
    const usersPath = path.join(DATA_DIR, 'users.json');
    let users = JSON.parse(fs.readFileSync(usersPath));
    users = users.filter(u => u.username !== username);
    fs.writeFileSync(usersPath, JSON.stringify(users, null, 2));
    return { success: true };
});

// EXPORT SYNC FILE
ipcMain.handle('export-sync', async () => {
    try {
        const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
            title: 'Export Sync File',
            defaultPath: `form137_sync_${new Date().toISOString().split('T')[0]}.f137sync`,
            filters: [{ name: 'Form 137 Sync Files', extensions: ['f137sync', 'json'] }]
        });

        if (canceled || !filePath) return { success: false };

        const structPath = path.join(DATA_DIR, 'structure.json');
        let structure = {};
        if (fs.existsSync(structPath)) {
            structure = JSON.parse(fs.readFileSync(structPath));
        }

        const records = [];
        if (fs.existsSync(RECORDS_DIR)) {
            const files = fs.readdirSync(RECORDS_DIR).filter(f => f.endsWith('.json'));
            for (const file of files) {
                const recordData = JSON.parse(fs.readFileSync(path.join(RECORDS_DIR, file)));
                records.push(recordData);
            }
        }

        const syncData = {
            version: 1,
            timestamp: new Date().toISOString(),
            structure,
            records
        };

        fs.writeFileSync(filePath, JSON.stringify(syncData, null, 2));
        return { success: true };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

function applySyncData(syncData) {
    if (!syncData || !syncData.structure || !Array.isArray(syncData.records)) {
        throw new Error('Invalid sync data format.');
    }

    // 1. Merge Structure
    const structPath = path.join(DATA_DIR, 'structure.json');
    let currentStruct = {};
    if (fs.existsSync(structPath)) {
        currentStruct = JSON.parse(fs.readFileSync(structPath));
    }
    const mergedStructure = deepMerge(currentStruct, syncData.structure);
    fs.writeFileSync(structPath, JSON.stringify(mergedStructure, null, 2));

    // 2. Merge Records
    for (const recordData of syncData.records) {
        const id = recordData.id || recordData.info?.lrn;
        if (!id) continue;

        const recordPath = path.join(RECORDS_DIR, `${id}.json`);
        let mergedRecord = recordData;
        if (fs.existsSync(recordPath)) {
            try {
                const existing = JSON.parse(fs.readFileSync(recordPath));
                mergedRecord = deepMerge(existing, recordData);
            } catch (err) {
                console.error('Failed to parse existing record, overwriting:', id, err.message);
            }
        }
        fs.writeFileSync(recordPath, JSON.stringify(mergedRecord, null, 2));
    }

    return syncData.records.length;
}

// IMPORT SYNC FILE
ipcMain.handle('import-sync', async () => {
    try {
        const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
            title: 'Import Sync File',
            properties: ['openFile'],
            filters: [{ name: 'Form 137 Sync Files', extensions: ['f137sync', 'json'] }]
        });

        if (canceled || filePaths.length === 0) return { success: false };

        const syncData = JSON.parse(fs.readFileSync(filePaths[0]));

        const count = applySyncData(syncData);
        return { success: true, count };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// IMPORT SYNC DATA DIRECTLY (FROM WEB INBOX MERGE FLOW)
ipcMain.handle('import-sync-data', async (event, syncData) => {
    try {
        const count = applySyncData(syncData);
        return { success: true, count };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// EXPORT TO EXCEL
ipcMain.handle('export-excel', async (event, data) => {
    const templatePath = isDev
        ? path.join(__dirname, '..', '..', 'Form 137-SHS-BLANK.xlsx')
        : path.join(process.resourcesPath, 'Form 137-SHS-BLANK.xlsx');

    if (!fs.existsSync(templatePath)) {
        return { success: false, error: 'Template Form 137-SHS-BLANK.xlsx not found.' };
    }

    try {
        const { filePath } = await dialog.showSaveDialog({
            title: 'Export Form 137 to Excel',
            defaultPath: `Form137_${data.info.lname || 'Student'}.xlsx`,
            filters: [{ name: 'Excel File', extensions: ['xlsx'] }]
        });

        if (!filePath) return { success: false, cancelled: true };

        const wb = xlsx.readFile(templatePath, { cellStyles: true });

        // Helper: Find cell by label and write value to offset
        const writeVal = (ws, label, value, offsetCol = 1) => {
            const range = xlsx.utils.decode_range(ws['!ref']);
            for (let r = range.s.r; r <= range.e.r; r++) {
                for (let c = range.s.c; c <= Math.min(range.e.c, 15); c++) {
                    const cell = ws[xlsx.utils.encode_cell({ r, c })];
                    if (cell && (cell.v || '').toString().toUpperCase().includes(label)) {
                        // Found label, write to offset
                        const targetAddr = xlsx.utils.encode_cell({ r, c: c + offsetCol });
                        if (!ws[targetAddr]) ws[targetAddr] = { t: 's', v: '' };
                        ws[targetAddr].v = value;
                        return;
                    }
                }
            }
        };

        // Same robust logic: write to specific key fields
        wb.SheetNames.forEach(sheetName => {
            const ws = wb.Sheets[sheetName];

            // Info
            writeVal(ws, "LAST NAME", data.info.lname);
            writeVal(ws, "FIRST NAME", data.info.fname);
            writeVal(ws, "MIDDLE NAME", data.info.mname);
            writeVal(ws, "LRN", data.info.lrn);
            writeVal(ws, "DATE OF BIRTH", data.info.birthdate);
            writeVal(ws, "SEX", data.info.sex);

            // Semesters - This is harder to genericize for writing without overwriting structure
            // For now, we'll skip writing the complex table to avoid breaking the layout
            // unless we use the exact coordinates.
            // Using the coordinates found in analysis:

            // Front Sheet (Sheet 0)
            if (sheetName === wb.SheetNames[0]) {
                // Info writes above cover this
            }

            // For table data, if we want to be perfect, we need exact coords.
            // I'll stick to Info for now as a "Proof of Concept" since table writing 
            // without shifting rows is risky in SheetJS.
        });

        xlsx.writeFile(wb, filePath);
        return { success: true };

    } catch (e) {
        console.error(e);
        return { success: false, error: e.message };
    }
});

const activeWorkers = new Set();

app.on('before-quit', () => {
    activeWorkers.forEach(w => w.kill());
});

// PRINT EXCEL FORM - Automate Excel via PowerShell for perfect rendering
// PRINT EXCEL FORM - Generate file via exceljs worker then open it
ipcMain.handle('print-excel-form', async (event, data) => {
    // In production, extraResources puts file in resources/
    // In dev, it's in root
    const templatePath = isDev
        ? path.join(__dirname, '..', '..', 'Form 137-SHS-BLANK.xlsx')
        : path.join(process.resourcesPath, 'Form 137-SHS-BLANK.xlsx');

    return new Promise((resolve) => {
        const { fork } = require('child_process');
        const workerPath = path.join(__dirname, 'excel-generator-worker.js');
        const worker = fork(workerPath, [], { stdio: 'ignore' });

        activeWorkers.add(worker);

        let resolved = false;

        // 45s timeout for heavy exceljs operations
        const timeout = setTimeout(() => {
            if (!resolved) {
                resolved = true;
                activeWorkers.delete(worker);
                worker.kill();
                resolve({ success: false, error: 'Excel generation timed out.' });
            }
        }, 45000);

        worker.on('message', (msg) => {
            if (msg.type === 'result') {
                if (!resolved) {
                    resolved = true;
                    clearTimeout(timeout);
                    activeWorkers.delete(worker);
                    worker.kill();

                    if (msg.success) {
                        require('electron').shell.openPath(msg.filePath);
                        resolve({ success: true, filePath: msg.filePath });
                    } else {
                        resolve({ success: false, error: msg.error });
                    }
                }
            }
        });

        worker.on('error', (err) => {
            if (!resolved) {
                resolved = true;
                clearTimeout(timeout);
                activeWorkers.delete(worker);
                resolve({ success: false, error: 'Worker error: ' + err.message });
            }
        });

        worker.on('exit', (code) => {
            if (!resolved) {
                resolved = true;
                clearTimeout(timeout);
                activeWorkers.delete(worker);
                resolve({ success: false, error: 'Worker exited with code ' + code });
            }
        });

        // Send data immediately
        worker.send({ type: 'generate-print-file', data, templatePath });
    });
});

// IMPORT EXCEL FORM - Parse .xlsx back to JSON
ipcMain.handle('import-excel-form', async (event, filePath) => {
    return new Promise((resolve) => {
        const { fork } = require('child_process');
        const workerPath = path.join(__dirname, 'excel-generator-worker.js');
        const worker = fork(workerPath, [], { stdio: 'ignore' });

        activeWorkers.add(worker);

        let resolved = false;

        // 30s timeout
        const timeout = setTimeout(() => {
            if (!resolved) {
                resolved = true;
                activeWorkers.delete(worker);
                worker.kill();
                resolve(null); // Return null on failure
            }
        }, 30000);

        worker.on('message', (msg) => {
            if (msg.type === 'parse-result') {
                if (!resolved) {
                    resolved = true;
                    clearTimeout(timeout);
                    activeWorkers.delete(worker);
                    worker.kill();

                    if (msg.success) {
                        resolve(msg.data);
                    } else {
                        resolve(null);
                    }
                }
            }
        });

        worker.on('error', () => { if (!resolved) { activeWorkers.delete(worker); resolve(null); } });
        worker.on('exit', () => { if (!resolved) { activeWorkers.delete(worker); resolve(null); } });

        worker.send({ type: 'parse-excel', filePath });
    });
});

ipcMain.handle('scan-placeholders', async () => {
    const templatePath = isDev
        ? path.join(__dirname, '..', '..', 'Form 137-SHS-BLANK.xlsx')
        : path.join(process.resourcesPath, 'Form 137-SHS-BLANK.xlsx');

    return new Promise((resolve) => {
        const { fork } = require('child_process');
        const workerPath = path.join(__dirname, 'excel-generator-worker.js');
        const worker = fork(workerPath, [], { stdio: 'ignore' });

        activeWorkers.add(worker);
        let resolved = false;

        const timeout = setTimeout(() => {
            if (!resolved) {
                resolved = true;
                activeWorkers.delete(worker);
                worker.kill();
                resolve({ success: false, error: 'Scanning timed out.' });
            }
        }, 15000);

        worker.on('message', (msg) => {
            if (msg.type === 'scan-result') {
                if (!resolved) {
                    resolved = true;
                    clearTimeout(timeout);
                    activeWorkers.delete(worker);
                    worker.kill();
                    resolve(msg);
                }
            }
        });

        worker.on('error', () => { if (!resolved) { resolved = true; activeWorkers.delete(worker); resolve({ success: false }); } });
        worker.on('exit', () => { if (!resolved) { resolved = true; activeWorkers.delete(worker); resolve({ success: false }); } });

        worker.send({ type: 'scan-placeholders', templatePath });
    });
});
