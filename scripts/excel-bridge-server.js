/* eslint-disable no-console */
console.log('--- Bridge Bootstrap Start ---');
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { exec, execFile } = require('child_process');

console.log('Loading dependencies...');
let ExcelJS;
try {
    ExcelJS = require('exceljs');
    console.log('ExcelJS loaded successfully.');
} catch (err) {
    console.error('FATAL: Failed to load exceljs:', err.message);
    process.exit(1);
}

const PORT = Number(process.env.BRIDGE_PORT || 8787);
const HOST = process.env.BRIDGE_HOST || '127.0.0.1';
const API_KEY = process.env.BRIDGE_API_KEY || '';
const ALLOW_ORIGIN = process.env.BRIDGE_ALLOW_ORIGIN || '*';

// Priority: 1. ENV, 2. Current Folder (for portable EXE), 3. public/ folder
const TEMPLATE_NAME = 'Form137_Template.xlsx';
const POSSIBLE_PATHS = [
    process.env.BRIDGE_TEMPLATE_PATH,
    path.join(process.cwd(), TEMPLATE_NAME),
    path.join(__dirname, TEMPLATE_NAME),
    path.join(process.cwd(), 'public', TEMPLATE_NAME),
    path.join(__dirname, '..', 'public', TEMPLATE_NAME)
];
const TEMPLATE_PATH = POSSIBLE_PATHS.find(p => p && fs.existsSync(p)) || POSSIBLE_PATHS[1];

const OUTPUT_DIR = process.env.BRIDGE_OUTPUT_DIR || path.join(os.tmpdir(), 'form137-exports');
const STARTUP_LOG_FILE = path.join(process.cwd(), 'excel-bridge-startup.log');

if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function logToFile(message) {
    try {
        fs.appendFileSync(
            STARTUP_LOG_FILE,
            `[${new Date().toISOString()}] ${message}\n`,
            'utf8'
        );
    } catch {
        // Ignore logging errors.
    }
}

function setCors(res) {
    res.setHeader('Access-Control-Allow-Origin', ALLOW_ORIGIN);
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type,X-Bridge-Key');
}

function sendJson(res, code, payload) {
    setCors(res);
    res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(payload));
}

function normalize(value) {
    if (value === undefined || value === null) return '';
    return String(value);
}

/**
 * Searches for a label text in the worksheet and writes to a cell relative to it.
 */
function writeByLabel(ws, label, value, offsetCol = 1, offsetRow = 0, maxCol = 24) {
    if (value === undefined || value === null || value === '') return false;
    if (!ws) return false;

    const search = label.toUpperCase();
    let found = false;

    ws.eachRow((row) => {
        if (found) return;
        row.eachCell((cell) => {
            if (found) return;
            const text = normalize(cell.value).toUpperCase();
            if (text.includes(search)) {
                // ExcelJS column/row are 1-indexed
                const targetRow = cell.row + offsetRow;
                const targetCol = cell.col + offsetCol;
                console.log(`Label [${label}] found at ${cell.address}. Writing to Col ${targetCol} Row ${targetRow}`);
                if (targetCol <= maxCol + offsetCol) {
                    ws.getCell(targetRow, targetCol).value = normalize(value);
                    found = true;
                }
            }
        });
    });
    return found;
}

function fillSemesterInfo(sheet, startRow, semData) {
    if (!sheet || !semData) return;

    // ExcelJS uses 1-based indexing for row/col
    const r = startRow;

    // School (Col 5 / E)
    sheet.getCell(r, 5).value = normalize(semData.school);
    // School ID (Col 29 / AC)
    sheet.getCell(r, 29).value = normalize(semData.schoolId);
    // Grade Level (Col 42 / AP)
    sheet.getCell(r, 42).value = normalize(semData.gradeLevel);
    // SY (Col 53 / BA)
    sheet.getCell(r, 53).value = normalize(semData.sy);
    // SEM (Col 63 / BK)
    sheet.getCell(r, 63).value = normalize(semData.semester);

    // Track/Strand (2 rows down, Col 7 / G)
    sheet.getCell(r + 2, 7).value = normalize(semData.trackStrand);
    // Section (Col 43 / AQ)
    sheet.getCell(r + 2, 43).value = normalize(semData.section);
}

function fillSemesterSubjects(sheet, startRow, subjects) {
    if (!sheet || !subjects || !Array.isArray(subjects)) return;

    subjects.forEach((subj, idx) => {
        const r = startRow + idx;
        if (r > 2000) return; // Safety

        // Col A (1): Type
        sheet.getCell(r, 1).value = normalize(subj.type);
        // Col I (9): Subject Name
        sheet.getCell(r, 9).value = normalize(subj.subject);
        // Col AT (46): Q1
        sheet.getCell(r, 46).value = normalize(subj.q1);
        // Col AY (51): Q2
        sheet.getCell(r, 51).value = normalize(subj.q2);
        // Col BD (56): Final
        sheet.getCell(r, 56).value = normalize(subj.final);
        // Col BI (61): Action
        sheet.getCell(r, 61).value = normalize(subj.action);
    });
}

/**
 * Fixes a common ExcelJS error: "Shared Formula master must exist above and or left of clone"
 * by converting shared formulas into regular formulas just before saving.
 */
function fixSharedFormulas(workbook) {
    workbook.worksheets.forEach(ws => {
        ws.eachRow(row => {
            row.eachCell(cell => {
                try {
                    // If it's a formula, we MUST either flatten it or strip it
                    if (cell.type === 6) {
                        // Try to get the values. If this crashes, the catch block handles it.
                        const f = cell.formula;
                        const res = cell.result;

                        if (f && typeof f === 'object' && f.shareType === 'shared') {
                            // Flatten shared formula to regular formula
                            cell.value = { formula: f.formula, result: res };
                        } else if (f && typeof f === 'object') {
                            // Ensure it's not a weird object that might cause issues later
                            cell.value = { formula: f.formula || String(f), result: res };
                        }
                    }
                } catch (e) {
                    // CRITICAL: If any part of the formula logic crashes, 
                    // we MUST strip the formula entirely to allow the file to save.
                    console.warn(`Nuking corrupted formula at ${cell.address}: ${e.message}`);
                    try {
                        // Attempt to keep the numeric/text result if possible
                        const fallback = cell.result;
                        cell.value = (fallback !== undefined && fallback !== null) ? fallback : '';
                    } catch (inner) {
                        cell.value = ''; // Total reset as last resort
                    }
                }
            });
        });
    });
}

function createSummaryWorkbook(data) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Form137');

    sheet.addRow(['FORM 137 - Student Permanent Record']);
    sheet.addRow([]);
    sheet.addRow(['Last Name', data.info?.lname || '']);
    sheet.addRow(['First Name', data.info?.fname || '']);
    sheet.addRow(['Middle Name', data.info?.mname || '']);
    sheet.addRow(['LRN', data.info?.lrn || '']);
    sheet.addRow(['Sex', data.info?.sex || '']);
    sheet.addRow(['Date of Birth', data.info?.birthdate || '']);
    sheet.addRow([]);

    const semesterKeys = ['semester1', 'semester2', 'semester3', 'semester4'];
    semesterKeys.forEach((semKey, index) => {
        const sem = data[semKey];
        if (!sem) return;

        sheet.addRow([`Semester ${index + 1}`]);
        sheet.addRow(['School', sem.school || '']);
        sheet.addRow(['School Year', sem.sy || '']);
        sheet.addRow(['Grade Level', sem.gradeLevel || '']);
        sheet.addRow(['Track/Strand', sem.trackStrand || '']);
        sheet.addRow(['Section', sem.section || '']);
        sheet.addRow(['Type', 'Subject', 'Q1', 'Q2', 'Final', 'Action']);

        (sem.subjects || []).forEach((subj) => {
            if (!subj || !subj.subject) return;
            sheet.addRow([
                subj.type || '',
                subj.subject || '',
                subj.q1 || '',
                subj.q2 || '',
                subj.final || '',
                subj.action || ''
            ]);
        });

        sheet.addRow(['General Average', sem.genAve || '']);
        sheet.addRow(['Remarks', sem.remarks || '']);
        sheet.addRow([]);
    });

    return workbook;
}

async function createWorkbookFromTemplate(data) {
    if (!fs.existsSync(TEMPLATE_PATH)) {
        return createSummaryWorkbook(data);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEMPLATE_PATH);

    const sheets = workbook.worksheets;
    const front = sheets.find(s => s.name.toUpperCase().includes('FRONT')) || sheets[0];
    const back = sheets.find(s => s.name.toUpperCase().includes('BACK')) || sheets[1] || front;

    if (front) {
        console.log('--- Front Sheet Structure (Rows 1-10) ---');
        for (let i = 1; i <= 10; i++) {
            const row = front.getRow(i);
            const vals = [];
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const v = normalize(cell.value);
                if (v) vals.push(`[${cell.address}]: "${v}"`);
            });
            if (vals.length) console.log(`Row ${i}: ${vals.join(' | ')}`);
        }
        writeByLabel(front, 'LAST NAME', data.info?.lname);
        writeByLabel(front, 'FIRST NAME', data.info?.fname);
        writeByLabel(front, 'MIDDLE NAME', data.info?.mname);
        writeByLabel(front, 'LRN', data.info?.lrn);
        writeByLabel(front, 'SEX', data.info?.sex);
        writeByLabel(front, 'DATE OF BIRTH', data.info?.birthdate, 3);
        writeByLabel(front, 'DATE OF SHS ADMISSION', data.info?.admissionDate, 5);

        writeByLabel(front, 'GEN. AVE', data.eligibility?.hsGenAve);
        writeByLabel(front, 'DATE OF GRADUATION', data.eligibility?.gradDate, 3);
        writeByLabel(front, 'NAME OF SCHOOL', data.eligibility?.schoolName, 2);
        writeByLabel(front, 'SCHOOL ADDRESS', data.eligibility?.schoolAddress, 1);

        // Sem 1 Start: Row 23, Subjects Start: Row 28
        fillSemesterInfo(front, 23, data.semester1);
        fillSemesterSubjects(front, 28, data.semester1?.subjects);

        // Sem 2 Start: Row 66, Subjects Start: Row 71
        fillSemesterInfo(front, 66, data.semester2);
        fillSemesterSubjects(front, 71, data.semester2?.subjects);
    }

    if (back) {
        writeByLabel(back, 'TRACK/STRAND', data.certification?.trackStrand, 3);
        writeByLabel(back, 'SHS GENERAL AVERAGE', data.certification?.genAve, 3);
        writeByLabel(back, 'DATE OF GRADUATION', data.certification?.gradDate, 2, 2);
        writeByLabel(back, 'NAME OF SCHOOL', data.certification?.schoolHead, 2);

        // Sem 3 Start: Row 4, Subjects Start: Row 11
        fillSemesterInfo(back, 4, data.semester3);
        fillSemesterSubjects(back, 11, data.semester3?.subjects);

        // Sem 4 Start: Row 46, Subjects Start: Row 51
        fillSemesterInfo(back, 46, data.semester4);
        fillSemesterSubjects(back, 51, data.semester4?.subjects);
    }

    return workbook;
}

function openFile(filePath) {
    return new Promise((resolve, reject) => {
        let command = '';
        if (process.platform === 'win32') {
            command = `cmd /c start "" "${filePath}"`;
        } else if (process.platform === 'darwin') {
            command = `open "${filePath}"`;
        } else {
            command = `xdg-open "${filePath}"`;
        }

        exec(command, (error) => {
            if (error) return reject(error);
            return resolve();
        });
    });
}

function printFileWindows(filePath) {
    return new Promise((resolve, reject) => {
        const safePath = String(filePath).replace(/'/g, "''");

        const psComScript = [
            "$ErrorActionPreference='Stop'",
            '$excel = New-Object -ComObject Excel.Application',
            '$excel.Visible = $false',
            '$excel.DisplayAlerts = $false',
            `$workbook = $excel.Workbooks.Open('${safePath}', 0, $true)`,
            '$workbook.PrintOut()',
            '$workbook.Close($false)',
            '$excel.Quit()',
            '[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null',
            '[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null'
        ].join('; ');

        const psFallbackScript = [
            "$ErrorActionPreference='Stop'",
            `Start-Process -FilePath '${safePath}' -Verb Print -PassThru | Wait-Process`
        ].join('; ');

        execFile(
            'powershell.exe',
            ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', psComScript],
            (error) => {
                if (!error) return resolve();

                execFile(
                    'powershell.exe',
                    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', psFallbackScript],
                    (fbError, fbStdout, fbStderr) => {
                        if (fbError) {
                            const details = (fbStderr || fbStdout || fbError.message || '').trim();
                            reject(new Error(`Failed to print.\nCOM Error: ${error.message}\nFallback Error: ${details}`));
                            return;
                        }
                        resolve();
                    }
                );
            }
        );
    });
}

function parseBody(req) {
    return new Promise((resolve, reject) => {
        let raw = '';
        req.on('data', chunk => {
            raw += chunk;
            if (raw.length > 5 * 1024 * 1024) {
                reject(new Error('Payload too large.'));
                req.destroy();
            }
        });
        req.on('end', () => {
            try {
                resolve(raw ? JSON.parse(raw) : {});
            } catch (err) {
                reject(new Error('Invalid JSON payload.'));
            }
        });
        req.on('error', reject);
    });
}

async function handleOpenExcel(req, res) {
    if (API_KEY) {
        const provided = req.headers['x-bridge-key'];
        if (!provided || provided !== API_KEY) {
            return sendJson(res, 401, { success: false, error: 'Unauthorized.' });
        }
    }

    try {
        console.log('Parsing request body...');
        const body = await parseBody(req);
        console.log('Body keys:', Object.keys(body || {}));
        const data = body && body.data;
        if (!data || !data.info) {
            console.error('Validation failed: Missing student data');
            return sendJson(res, 400, { success: false, error: 'Missing student data.' });
        }

        const autoPrint = Boolean(body && body.autoPrint);
        const openAfterPrint = body && body.openAfterPrint !== false;

        console.log(`Generating workbook using template: ${TEMPLATE_PATH}`);
        if (!fs.existsSync(TEMPLATE_PATH)) {
            console.warn('Template not found! Falling back to summary sheet.');
        }

        const workbook = await createWorkbookFromTemplate(data);
        const lastName = (data.info?.lname || 'Student').replace(/[^a-z0-9]/gi, '_');
        const filePath = path.join(OUTPUT_DIR, `Form137_${lastName}_${Date.now()}.xlsx`);
        console.log(`Saving to: ${filePath}`);

        // IMPORTANT: Fix Shared Formula bug before saving
        fixSharedFormulas(workbook);

        await workbook.xlsx.writeFile(filePath);

        let warning = '';
        let printed = false;

        if (autoPrint) {
            if (process.platform === 'win32') {
                try {
                    await printFileWindows(filePath);
                    printed = true;
                } catch (err) {
                    warning = `Auto print failed: ${err.message}`;
                    console.warn(warning);
                }
            } else {
                warning = 'Auto print is only supported on Windows bridge hosts.';
                console.warn(warning);
            }

            if (openAfterPrint) {
                await openFile(filePath);
            }
        } else {
            await openFile(filePath);
        }

        return sendJson(res, 200, { success: true, filePath, printed, warning });
    } catch (err) {
        console.error('Bridge request error:', err.stack || err);
        logToFile(`Request error: ${err.stack || err.message}`);
        return sendJson(res, 500, { success: false, error: err.message });
    }
}

const server = http.createServer(async (req, res) => {
    try {
        if (req.method === 'OPTIONS') {
            setCors(res);
            res.writeHead(204);
            res.end();
            return;
        }

        if (req.method === 'GET' && req.url === '/health') {
            return sendJson(res, 200, {
                success: true,
                service: 'excel-bridge',
                host: HOST,
                port: PORT
            });
        }

        if (req.method === 'POST' && req.url === '/open-excel') {
            await handleOpenExcel(req, res);
            return;
        }

        return sendJson(res, 404, { success: false, error: 'Not found.' });
    } catch (err) {
        console.error('Server error:', err);
        return sendJson(res, 500, { success: false, error: err.message });
    }
});

function getStartupHint(error) {
    if (!error || !error.code) return '';
    if (error.code === 'EADDRINUSE') {
        return `Port ${PORT} is already in use. Close the app using that port, or start bridge with BRIDGE_PORT=8788.`;
    }
    if (error.code === 'EACCES') {
        return `No permission to bind ${HOST}:${PORT}. Try a non-privileged port like 8787 or 8788.`;
    }
    return '';
}

server.on('error', (error) => {
    const hint = getStartupHint(error);
    const detail = `Bridge failed to start on ${HOST}:${PORT} (${error.code || 'UNKNOWN'}): ${error.message}`;

    console.error(detail);
    if (hint) console.error(hint);
    console.error(`Log file: ${STARTUP_LOG_FILE}`);

    logToFile(detail);
    if (hint) logToFile(`Hint: ${hint}`);

    // CRITICAL: Exit so the launcher knows it failed
    process.exit(1);
});

process.on('uncaughtException', (error) => {
    const detail = `Uncaught exception: ${error.stack || error.message}`;
    console.error(detail);
    console.error(`Log file: ${STARTUP_LOG_FILE}`);
    logToFile(detail);
    process.exit(1);
});

process.on('unhandledRejection', (reason) => {
    const detail = `Unhandled rejection: ${reason && reason.stack ? reason.stack : String(reason)}`;
    console.error(detail);
    console.error(`Log file: ${STARTUP_LOG_FILE}`);
    logToFile(detail);
    process.exit(1);
});

console.log('Starting server listener...');
server.listen(PORT, HOST, () => {
    console.log('================================================');
    console.log('       SF10 EXCEL BRIDGE SERVER RUNNING         ');
    console.log('================================================');
    console.log(`Status:  Online`);
    console.log(`Address: http://${HOST}:${PORT}`);
    console.log(`URL:     ${HOST === '127.0.0.1' ? 'Local' : HOST}`);
    console.log(`Auth:    ${API_KEY ? `Enabled (Key: ${API_KEY})` : 'Disabled'}`);
    console.log('------------------------------------------------');

    if (fs.existsSync(TEMPLATE_PATH)) {
        console.log(`Template: [FOUND] ${path.basename(TEMPLATE_PATH)}`);
    } else {
        console.log(`Template: [MISSING] Fallback to simple list mode.`);
        console.warn(`Looking for: ${TEMPLATE_NAME}`);
    }

    console.log(`Outputs:  ${OUTPUT_DIR}`);
    console.log('================================================');
    console.log('Keep this window open while using the website.');
    console.log('Press Ctrl+C to close.');
});
