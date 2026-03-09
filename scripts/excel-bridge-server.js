/* eslint-disable no-console */
console.log('--- Bridge Bootstrap Start ---');
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { exec, execFile } = require('child_process');

function waitForEnter(exitCode = 1) {
    console.log('\nPress Enter to exit...');
    const rl = require('readline').createInterface({ input: process.stdin, output: process.stdout });
    rl.on('line', () => process.exit(exitCode));
}

const INSTALL_DIR = 'C:\\Form137Bridge';
const STARTUP_LOG_FILE = path.join(INSTALL_DIR, 'excel-bridge-startup.log');
function logToFile(message) {
    try {
        if (!fs.existsSync(INSTALL_DIR)) fs.mkdirSync(INSTALL_DIR, { recursive: true });
        fs.appendFileSync(STARTUP_LOG_FILE, `[${new Date().toISOString()}] ${message}\n`, 'utf8');
    } catch { /* ignore */ }
}

process.on('uncaughtException', (error) => {
    const detail = `Uncaught exception: ${error.stack || error.message}`;
    console.error(detail);
    logToFile(detail);
    waitForEnter(1);
});

process.on('unhandledRejection', (reason) => {
    const detail = `Unhandled rejection: ${reason && reason.stack ? reason.stack : String(reason)}`;
    console.error(detail);
    logToFile(detail);
    waitForEnter(1);
});

let ExcelJS;
const API_KEY = process.env.BRIDGE_API_KEY || 'sf10-bridge-2024';
try {
    console.log('Loading dependencies...');
    ExcelJS = require('exceljs');
    console.log('ExcelJS loaded successfully.');
} catch (err) {
    console.error('Failed to load dependencies:', err);
    logToFile(`Dependency load error: ${err.message}`);
}

try {
    const PORT = Number(process.env.BRIDGE_PORT || 8787);
    const HOST = process.env.BRIDGE_HOST || '0.0.0.0';
    const ALLOW_ORIGIN = process.env.BRIDGE_ALLOW_ORIGIN || '*';

    const TEMPLATE_CANDIDATE_NAMES = ['Form137_Template.xlsx', 'Form 137-SHS-BLANK.xlsx', 'PLACEHOLDER(ALL).xlsx'];
    const POSSIBLE_PATHS = [
        process.env.BRIDGE_TEMPLATE_PATH,
        path.join(INSTALL_DIR, TEMPLATE_CANDIDATE_NAMES[0]),
        path.join(INSTALL_DIR, TEMPLATE_CANDIDATE_NAMES[1]),
        ...TEMPLATE_CANDIDATE_NAMES.flatMap((name) => ([
            path.join(process.cwd(), name),
            path.join(__dirname, name),
            path.join(process.cwd(), 'public', name),
            path.join(__dirname, '..', 'public', name),
            path.join(process.cwd(), 'release', name),
            path.join(__dirname, '..', 'release', name)
        ]))
    ];
    const TEMPLATE_PATH = POSSIBLE_PATHS.find(p => p && fs.existsSync(p))
        || path.join(INSTALL_DIR, TEMPLATE_CANDIDATE_NAMES[0]);

    if (!fs.existsSync(TEMPLATE_PATH)) {
        console.warn(`[CRITICAL] Template not found in any possible path. Bridge will use fallback generator.`);
        logToFile(`[CRITICAL] Template not found: ${TEMPLATE_PATH}`);
    }

    const OUTPUT_DIR = process.env.BRIDGE_OUTPUT_DIR || path.join(INSTALL_DIR, 'exports');

    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }

    function setCors(req, res) {
        if (!res) return;
        const origin = (req && req.headers && req.headers.origin) || ALLOW_ORIGIN;
        res.setHeader('Access-Control-Allow-Origin', origin);
        res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type,X-Bridge-Key,ngrok-skip-browser-warning');
        res.setHeader('Access-Control-Allow-Credentials', 'true');
    }

    function sendJson(req, res, code, payload) {
        setCors(req, res);
        res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify(payload));
    }

    function normalize(value) {
        if (value === undefined || value === null) return '';
        return String(value);
    }

    function findMergeBounds(ws, row, col) {
        if (!ws) return null;
        try {
            const cell = ws.getCell(row, col);
            const address = cell.address;
            const internalMerges = ws._merges;
            if (internalMerges) {
                if (cell && cell.isMerged && cell.master && cell.master.address !== address) {
                    const m = cell.master;
                    return { startCol: m.col, startRow: m.row, endCol: m.col, endRow: m.row };
                }
                for (const mKey in internalMerges) {
                    const m = internalMerges[mKey];
                    if (m && row >= m.top && row <= m.bottom && col >= m.left && col <= m.right) {
                        return { startCol: m.left, startRow: m.top, endCol: m.right, endRow: m.bottom };
                    }
                }
            }
            const modelMerges = ws.model && ws.model.merges;
            if (Array.isArray(modelMerges)) {
                for (const ref of modelMerges) {
                    const [start, end] = ref.split(':');
                    const sCell = ws.getCell(start);
                    const eCell = ws.getCell(end);
                    if (row >= sCell.row && row <= eCell.row && col >= sCell.col && col <= eCell.col) {
                        return { startCol: sCell.col, startRow: sCell.row, endCol: eCell.col, endRow: eCell.row };
                    }
                }
            }
        } catch (e) {
            logToFile(`Merge check error at ${row},${col}: ${e.message}`);
        }
        return null;
    }

    function writeToCellMergeAware(ws, row, col, value) {
        if (!ws || row < 1 || col < 1) return;
        const merge = findMergeBounds(ws, row, col);
        const targetCell = merge ? ws.getCell(merge.startRow, merge.startCol) : ws.getCell(row, col);
        targetCell.value = normalize(value);
    }

    function excelCellToString(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'string') return value;
        if (typeof value === 'number' || typeof value === 'boolean') return String(value);
        if (value instanceof Date) return value.toISOString();
        if (typeof value === 'object') {
            if (Array.isArray(value.richText)) return value.richText.map(part => part.text || '').join('');
            if (typeof value.text === 'string') return value.text;
            if (value.result !== undefined && value.result !== null) return excelCellToString(value.result);
        }
        return String(value);
    }

    function findLabelPos(ws, label, rowRange) {
        if (!ws) return null;
        const search = label.toUpperCase();
        for (let r = rowRange.start; r <= rowRange.end; r++) {
            const row = ws.getRow(r);
            if (!row) continue;
            let found = null;
            row.eachCell({ includeEmpty: false }, (cell) => {
                if (found) return;
                const text = normalize(cell.value).toUpperCase();
                if (text.includes(search)) {
                    found = { row: cell.row, col: cell.col };
                }
            });
            if (found) return found;
        }
        return null;
    }

    function writeByLabel(ws, label, value, offsetCol = 1, offsetRow = 0, anchorFromEnd = true, rowRange = { start: 1, end: 1000 }) {
        if (value === undefined || value === null || value === '') return false;
        const pos = findLabelPos(ws, label, rowRange);
        if (!pos) {
            logToFile(`[WARN] Label [${label}] NOT FOUND in rows ${rowRange.start}-${rowRange.end}`);
            return false;
        }
        const merge = findMergeBounds(ws, pos.row, pos.col);
        const baseCol = (anchorFromEnd && merge) ? merge.endCol : pos.col;
        writeToCellMergeAware(ws, pos.row + offsetRow, baseCol + offsetCol, value);
        logToFile(`[SUCCESS] Label [${label}] found. Wrote [${value}] to offset ${offsetRow},${offsetCol}`);
        return true;
    }

    function buildPlaceholderMap(data) {
        const info = data?.info || {};
        const eligibility = data?.eligibility || {};
        const certification = data?.certification || {};
        const semester1 = data?.semester1 || {};
        const semester2 = data?.semester2 || {};
        const semester3 = data?.semester3 || {};
        const semester4 = data?.semester4 || {};

        const placeholderMap = {
            lname: normalize(info.lname),
            fname: normalize(info.fname),
            mname: normalize(info.mname),
            lrn: normalize(info.lrn),
            sex: normalize(info.sex),
            birthdate: normalize(info.birthdate),
            admission_date: normalize(info.admissionDate),
            hs_school: normalize(eligibility.schoolName),
            hs_addr: normalize(eligibility.schoolAddress),
            grad_date: normalize(eligibility.gradDate || certification.gradDate),
            cert_date: normalize(certification.certDate),
            cert_remarks: normalize(certification.remarks),
            date_issued: normalize(certification.dateIssued),
            hs_check: eligibility.hsCompleter ? 'X' : '',
            jhs_check: eligibility.jhsCompleter ? 'X' : '',
            pept_check: eligibility.pept ? 'X' : '',
            als_check: eligibility.als ? 'X' : '',
            others_check: eligibility.others ? 'X' : '',
            hs_ave: normalize(eligibility.hsGenAve),
            jhs_ave: normalize(eligibility.jhsGenAve),
            pept_rating: normalize(eligibility.peptRating),
            als_rating: normalize(eligibility.alsRating),
            exam_date: normalize(eligibility.examDate),
            clc_name: normalize(eligibility.clcName),
            others_spec: normalize(eligibility.othersSpec),
            lastname: normalize(info.lname),
            firstname: normalize(info.fname),
            middlename: normalize(info.mname),
            LRN: normalize(info.lrn),
            'date of birth': normalize(info.birthdate),
            dateofadmission: normalize(info.admissionDate),
            genave: normalize(semester1.genAve),
            track: normalize(certification.trackStrand || semester1.trackStrand),
            strand: normalize(certification.trackStrand || semester1.trackStrand),
            section: normalize(semester1.section),
            shs_gen_ave: normalize(certification.genAve),
            awards: normalize(certification.awards),
            school_head: normalize(certification.schoolHead)
        };

        const semesters = [semester1, semester2, semester3, semester4];
        semesters.forEach((sem, idx) => {
            const sNum = idx + 1;
            placeholderMap[`s${sNum}_school`] = normalize(sem.school);
            placeholderMap[`s${sNum}_id`] = normalize(sem.schoolId);
            placeholderMap[`s${sNum}_level`] = normalize(sem.gradeLevel);
            placeholderMap[`s${sNum}_sy`] = normalize(sem.sy);
            placeholderMap[`s${sNum}_sem`] = normalize(sem.semester || sem.sem);
            placeholderMap[`s${sNum}_ave`] = normalize(sem.genAve);
            placeholderMap[`s${sNum}_track`] = normalize(sem.trackStrand);
            placeholderMap[`s${sNum}_section`] = normalize(sem.section);

            (sem.subjects || []).forEach((subj, subjIdx) => {
                const rowNum = subjIdx + 1;
                placeholderMap[`s${sNum}sub_${rowNum}`] = normalize(subj?.subject);
                placeholderMap[`s${sNum}q1_${rowNum}`] = normalize(subj?.q1);
                placeholderMap[`s${sNum}q2_${rowNum}`] = normalize(subj?.q2);
                placeholderMap[`s${sNum}fin_${rowNum}`] = normalize(subj?.final);
                placeholderMap[`s${sNum}act_${rowNum}`] = normalize(subj?.action || subj?.actionTaken);
            });
        });
        return placeholderMap;
    }

    function replacePlaceholders(workbook, data) {
        const placeholderMap = buildPlaceholderMap(data);
        const keys = Object.keys(placeholderMap);
        if (!keys.length) return;
        workbook.worksheets.forEach((ws) => {
            ws.eachRow((row) => {
                row.eachCell((cell) => {
                    const text = excelCellToString(cell.value);
                    if (!text || !text.includes('%(')) return;
                    let replaced = text;
                    keys.forEach((key) => {
                        const token = `%(${key})`;
                        if (replaced.includes(token)) {
                            replaced = replaced.split(token).join(normalize(placeholderMap[key]));
                        }
                    });
                    if (replaced !== text) {
                        cell.value = replaced;
                    }
                });
            });
        });
    }

    function fillSemesterInfo(sheet, startRow, semData) {
        if (!sheet || !semData) return;
        const rowRange = { start: startRow, end: startRow + 8 };
        writeByLabel(sheet, 'SCHOOL:', semData.school, 1, 0, true, rowRange);
        writeByLabel(sheet, 'SCHOOL ID:', semData.schoolId, 1, 0, true, rowRange);
        writeByLabel(sheet, 'GRADE LEVEL:', semData.gradeLevel, 1, 0, true, rowRange);
        writeByLabel(sheet, 'SY:', semData.sy, 1, 0, true, rowRange);
        writeByLabel(sheet, 'SEM:', semData.semester || semData.sem, 1, 0, true, rowRange);
        writeByLabel(sheet, 'TRACK/STRAND', semData.trackStrand, 1, 0, true, rowRange);
        writeByLabel(sheet, 'SECTION', semData.section, 1, 0, true, rowRange);
    }

    function fillSemesterSubjects(sheet, startRow, subjects) {
        if (!sheet || !subjects || !Array.isArray(subjects)) return;

        const typeMap = {
            'CORE': 'Core',
            'APPLIED': 'Applied',
            'SPECIALIZED': 'Specialized',
            'SPE': 'Specialized',
            'SPEC': 'Specialized',
            'APP': 'Applied'
        };

        subjects.forEach((subj, idx) => {
            const r = startRow + idx;
            if (r > 2000) return;

            // Map type or fallback to index
            let typeVal = subj.type || '';
            const upperType = typeVal.toUpperCase();
            if (typeMap[upperType]) {
                typeVal = typeMap[upperType];
            } else if (!typeVal) {
                typeVal = idx + 1;
            }

            // Col A (1): Type
            writeToCellMergeAware(sheet, r, 1, typeVal);
            // Col E (5): Subject Name
            writeToCellMergeAware(sheet, r, 5, subj.subject);
            // Col AT (46): Q1
            writeToCellMergeAware(sheet, r, 46, subj.q1);
            // Col AY (51): Q2
            writeToCellMergeAware(sheet, r, 51, subj.q2);
            // Col BD (56): Final
            writeToCellMergeAware(sheet, r, 56, subj.final);
            // Col BI (61): Action
            writeToCellMergeAware(sheet, r, 61, subj.action || subj.actionTaken);
        });
    }

    function fillRemedialInfo(sheet, startRow, remedial) {
        if (!sheet || !remedial) return;
        (remedial.subjects || []).forEach((subj, idx) => {
            const r = startRow + idx;
            writeToCellMergeAware(sheet, r, 9, subj.subject);
            writeToCellMergeAware(sheet, r, 56, subj.semGrade);
            writeToCellMergeAware(sheet, r, 61, subj.remedialMark);
            writeToCellMergeAware(sheet, r, 66, subj.recomputedGrade);
            writeToCellMergeAware(sheet, r, 71, subj.action);
        });
    }

    function fixSharedFormulas(workbook) {
        workbook.worksheets.forEach(ws => {
            ws.eachRow(row => {
                row.eachCell(cell => {
                    try {
                        if (cell.type === 6) {
                            const f = cell.formula;
                            const res = cell.result;
                            if (f && typeof f === 'object' && f.shareType === 'shared') {
                                cell.value = { formula: f.formula, result: res };
                            } else if (f && typeof f === 'object') {
                                cell.value = { formula: f.formula || String(f), result: res };
                            }
                        }
                    } catch (e) {
                        console.warn(`Nuking corrupted formula at ${cell.address}: ${e.message}`);
                        try {
                            const fallback = cell.result;
                            cell.value = (fallback !== undefined && fallback !== null) ? fallback : '';
                        } catch (inner) {
                            cell.value = '';
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
        const semesterKeys = ['semester1', 'semester2', 'semester3', 'semester4'];
        semesterKeys.forEach((semKey, index) => {
            const sem = data[semKey];
            if (!sem) return;
            sheet.addRow([`Semester ${index + 1}`]);
            sheet.addRow(['School', sem.school || '']);
            (sem.subjects || []).forEach((subj) => {
                if (!subj || !subj.subject) return;
                sheet.addRow([subj.type || '', subj.subject || '', subj.q1 || '', subj.q2 || '', subj.final || '', subj.action || '']);
            });
            sheet.addRow(['General Average', sem.genAve || '']);
            sheet.addRow([]);
        });
        return workbook;
    }

    async function createWorkbookFromTemplate(data) {
        if (!fs.existsSync(TEMPLATE_PATH)) return createSummaryWorkbook(data);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(TEMPLATE_PATH);
        replacePlaceholders(workbook, data);
        const sheets = workbook.worksheets;
        const front = sheets.find(s => s.name.toUpperCase().includes('FRONT')) || sheets[0];
        const back = sheets.find(s => s.name.toUpperCase().includes('BACK')) || sheets[1] || front;
        if (front) {
            writeByLabel(front, 'LAST NAME', data.info?.lname);
            writeByLabel(front, 'FIRST NAME', data.info?.fname);
            writeByLabel(front, 'MIDDLE NAME', data.info?.mname);
            writeByLabel(front, 'LRN', data.info?.lrn);
            writeByLabel(front, 'SEX', data.info?.sex);
            writeByLabel(front, 'DATE OF BIRTH', data.info?.birthdate, 1);
            writeByLabel(front, 'DATE OF SHS ADMISSION', data.info?.admissionDate, 1);
            if (data.eligibility?.hsCompleter) writeByLabel(front, 'High School Completer*', 'X', -1, 0, false);
            if (data.eligibility?.jhsCompleter) writeByLabel(front, 'Junior High School Completer', 'X', -1, 0, false);
            if (data.eligibility?.pept) writeByLabel(front, 'PEPT Passer**', 'X', -1, 0, false);
            if (data.eligibility?.als) writeByLabel(front, 'ALS A&E Passer***', 'X', -1, 0, false);
            if (data.eligibility?.others) writeByLabel(front, 'Others (Pls. Specify):', 'X', -1, 0, false);
            writeByLabel(front, 'Gen. Ave', data.eligibility?.hsGenAve || data.eligibility?.jhsGenAve);
            writeByLabel(front, 'DATE OF GRADUATION', data.eligibility?.gradDate, 3);
            writeByLabel(front, 'NAME OF SCHOOL', data.eligibility?.schoolName, 2);
            writeByLabel(front, 'SCHOOL ADDRESS', data.eligibility?.schoolAddress, 1);
            writeByLabel(front, 'Rating:', data.eligibility?.peptRating || data.eligibility?.alsRating);
            writeByLabel(front, 'Date of Examination/Assessment', data.eligibility?.examDate);
            writeByLabel(front, 'Name/Address of Community Learning Center', data.eligibility?.clcName);
            writeByLabel(front, 'Specify):', data.eligibility?.othersSpec);

            // Sem 1
            fillSemesterInfo(front, 23, data.semester1);
            const s1Table = findLabelPos(front, 'SUBJECTS', { start: 23, end: 40 });
            if (s1Table) fillSemesterSubjects(front, s1Table.row + 2, data.semester1?.subjects);

            // Sem 2
            const s2Start = findLabelPos(front, 'SCHOOL:', { start: 40, end: 120 });
            if (s2Start) {
                fillSemesterInfo(front, s2Start.row, data.semester2);
                const s2Table = findLabelPos(front, 'SUBJECTS', { start: s2Start.row, end: s2Start.row + 20 });
                if (s2Table) fillSemesterSubjects(front, s2Table.row + 2, data.semester2?.subjects);
            }
        }
        if (back) {
            writeByLabel(back, 'TRACK/STRAND', data.certification?.trackStrand, 3);
            writeByLabel(back, 'SHS GENERAL AVERAGE', data.certification?.genAve, 3);
            writeByLabel(back, 'DATE OF GRADUATION', data.certification?.gradDate, 2, 2);
            writeByLabel(back, 'NAME OF SCHOOL', data.certification?.schoolHead, 2);
            const s3Start = findLabelPos(back, 'SCHOOL:', { start: 1, end: 30 });
            if (s3Start) {
                fillSemesterInfo(back, s3Start.row, data.semester3);
                const s3Table = findLabelPos(back, 'SUBJECTS', { start: s3Start.row, end: s3Start.row + 20 });
                if (s3Table) fillSemesterSubjects(back, s3Table.row + 2, data.semester3?.subjects);
            }
            const s4Start = findLabelPos(back, 'SCHOOL:', { start: 30, end: 150 });
            if (s4Start) {
                fillSemesterInfo(back, s4Start.row, data.semester4);
                const s4Table = findLabelPos(back, 'SUBJECTS', { start: s4Start.row, end: s4Start.row + 20 });
                if (s4Table) fillSemesterSubjects(back, s4Table.row + 2, data.semester4?.subjects);
            }
        }
        return workbook;
    }

    function openFile(filePath) {
        return new Promise((resolve, reject) => {
            let command = (process.platform === 'win32') ? `cmd /c start "" "${filePath}"` : (process.platform === 'darwin' ? `open "${filePath}"` : `xdg-open "${filePath}"`);
            exec(command, (error) => error ? reject(error) : resolve());
        });
    }

    function printFileWindows(filePath) {
        return new Promise((resolve, reject) => {
            const safePath = String(filePath).replace(/'/g, "''");
            const psComScript = ["$ErrorActionPreference='Stop'", '$excel = New-Object -ComObject Excel.Application', '$excel.Visible = $false', '$excel.DisplayAlerts = $false', `$workbook = $excel.Workbooks.Open('${safePath}', 0, $true)`, '$workbook.PrintOut()', '$workbook.Close($false)', '$excel.Quit()', '[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null', '[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null'].join('; ');
            const psFallbackScript = ["$ErrorActionPreference='Stop'", `Start-Process -FilePath '${safePath}' -Verb Print -PassThru | Wait-Process`].join('; ');
            execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', psComScript], (error) => {
                if (!error) return resolve();
                execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', psFallbackScript], (fbError, fbStdout, fbStderr) => fbError ? reject(new Error(`Failed to print.\nCOM Error: ${error.message}\nFallback Error: ${(fbStderr || fbStdout || fbError.message || '').trim()}`)) : resolve());
            });
        });
    }

    function parseBody(req) {
        return new Promise((resolve, reject) => {
            let raw = '';
            req.on('data', chunk => {
                raw += chunk;
                if (raw.length > 5 * 1024 * 1024) { req.destroy(); reject(new Error('Payload too large.')); }
            });
            req.on('end', () => { try { resolve(raw ? JSON.parse(raw) : {}); } catch (err) { reject(new Error('Invalid JSON payload.')); } });
            req.on('error', reject);
        });
    }

    async function handleOpenExcel(req, res) {
        if (API_KEY) {
            const provided = req.headers['x-bridge-key'];
            if (!provided || provided !== API_KEY) return sendJson(req, res, 401, { success: false, error: 'Unauthorized.' });
        }
        try {
            const body = await parseBody(req);
            const data = body && body.data;
            if (!data || !data.info) return sendJson(req, res, 400, { success: false, error: 'Missing student data.' });
            console.log('--- DATA RECEIVED ---');
            console.log(`Student: ${data.info.fname} ${data.info.lname}`);
            const workbook = await createWorkbookFromTemplate(data);
            const lastName = (data.info?.lname || 'Student').replace(/[^a-z0-9]/gi, '_');
            const filePath = path.join(OUTPUT_DIR, `Form137_${lastName}_${Date.now()}.xlsx`);
            fixSharedFormulas(workbook);
            await workbook.xlsx.writeFile(filePath);
            let printed = false;
            if (body && body.autoPrint && process.platform === 'win32') {
                try { await printFileWindows(filePath); printed = true; } catch (err) { console.warn(`Print failed: ${err.message}`); }
            }
            if (!body || body.openAfterPrint !== false) await openFile(filePath);
            if (body && body.returnFile) {
                res.writeHead(200, {
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'Content-Disposition': `attachment; filename="Form137_${lastName}.xlsx"`
                });
                return fs.createReadStream(filePath).pipe(res);
            }
            return sendJson(req, res, 200, { success: true, filePath, printed });
        } catch (err) {
            console.error('Bridge error:', err);
            logToFile(`Request error: ${err.stack || err.message}`);
            return sendJson(req, res, 500, { success: false, error: err.message });
        }
    }

    const server = http.createServer(async (req, res) => {
        if (req.method === 'OPTIONS') { setCors(req, res); res.writeHead(204); res.end(); return; }
        if (req.method === 'GET' && req.url === '/health') return sendJson(req, res, 200, { success: true, service: 'excel-bridge', host: HOST, port: PORT });

        // --- Download Routes ---
        if (req.method === 'GET' && req.url === '/download/installer') {
            const target = path.join(process.cwd(), 'scripts', 'setup-portable-bridge.ps1');
            console.log(`Download Installer: ${target} (exists: ${fs.existsSync(target)})`);
            if (fs.existsSync(target)) {
                res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Content-Disposition': 'attachment; filename="setup-portable-bridge.ps1"' });
                return fs.createReadStream(target).pipe(res);
            }
            return sendJson(req, res, 404, { success: false, error: 'Installer script not found.' });
        }

        if (req.method === 'GET' && req.url === '/download/exe') {
            const target = path.join(process.cwd(), 'release', 'excel-bridge-server.exe');
            if (fs.existsSync(target)) {
                res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Content-Disposition': 'attachment; filename="excel-bridge-server.exe"' });
                return fs.createReadStream(target).pipe(res);
            }
            return sendJson(req, res, 404, { success: false, error: 'Bridge executable not found.' });
        }

        if (req.method === 'GET' && req.url === '/download/template') {
            if (fs.existsSync(TEMPLATE_PATH)) {
                res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="${path.basename(TEMPLATE_PATH)}"` });
                return fs.createReadStream(TEMPLATE_PATH).pipe(res);
            }
            return sendJson(req, res, 404, { success: false, error: 'Template file not found.' });
        }

        if (req.method === 'POST' && req.url === '/open-excel') { await handleOpenExcel(req, res); return; }
        return sendJson(req, res, 404, { success: false, error: 'Not found.' });
    });

    server.on('error', (error) => {
        const detail = `Bridge failed on ${HOST}:${PORT}: ${error.message}`;
        console.error(detail);
        logToFile(detail);
        waitForEnter(1);
    });

    server.listen(PORT, HOST, () => {
        console.log(`SF10 EXCEL BRIDGE ONLINE: http://${HOST}:${PORT}`);
        console.log(`Template: ${TEMPLATE_PATH}`);
    });

} catch (err) {
    console.error(`FATAL: ${err.stack || err}`);
    waitForEnter(1);
}
