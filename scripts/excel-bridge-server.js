/* eslint-disable no-console */
console.log('--- Bridge Bootstrap Start ---');
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { exec, execFile } = require('child_process');

let ExcelJS;
const API_KEY = process.env.BRIDGE_API_KEY || 'sf10-bridge-2024';
try {
    console.log('Loading dependencies...');
    ExcelJS = require('exceljs');
    console.log('ExcelJS loaded successfully.');

    const PORT = Number(process.env.BRIDGE_PORT || 8787);
    const HOST = process.env.BRIDGE_HOST || '127.0.0.1';
    const ALLOW_ORIGIN = process.env.BRIDGE_ALLOW_ORIGIN || '*';

    // Priority: 1. ENV, 2. Placeholder template, 3. Legacy template
    const TEMPLATE_CANDIDATE_NAMES = ['PLACEHOLDER(ALL).xlsx', 'Form137_Template.xlsx'];
    const POSSIBLE_PATHS = [
        process.env.BRIDGE_TEMPLATE_PATH,
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
        || path.join(process.cwd(), TEMPLATE_CANDIDATE_NAMES[1]);

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

    function setCors(req, res) {
        const origin = req.headers.origin || ALLOW_ORIGIN;
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
        if (!ws || !ws.model || !Array.isArray(ws.model.merges)) return null;
        for (const ref of ws.model.merges) {
            const [start, end] = ref.split(':');
            const startCell = ws.getCell(start);
            const endCell = ws.getCell(end);
            if (
                row >= startCell.row
                && row <= endCell.row
                && col >= startCell.col
                && col <= endCell.col
            ) {
                return {
                    startCol: startCell.col,
                    endCol: endCell.col,
                    startRow: startCell.row,
                    endRow: endCell.row
                };
            }
        }
        return null;
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

            // Eligibility Checkboxes (returns 'X' if true)
            hs_check: eligibility.hsCompleter ? 'X' : '',
            jhs_check: eligibility.jhsCompleter ? 'X' : '',
            pept_check: eligibility.pept ? 'X' : '',
            als_check: eligibility.als ? 'X' : '',
            others_check: eligibility.others ? 'X' : '',

            // Eligibility Ratings/Dates
            hs_ave: normalize(eligibility.hsGenAve),
            jhs_ave: normalize(eligibility.jhsGenAve),
            pept_rating: normalize(eligibility.peptRating),
            als_rating: normalize(eligibility.alsRating),
            exam_date: normalize(eligibility.examDate),
            clc_name: normalize(eligibility.clcName),
            others_spec: normalize(eligibility.othersSpec),

            // Compatibility aliases for existing templates
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
                placeholderMap[`s${sNum}act_${rowNum}`] = normalize(subj?.action);
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

    /**
     * Searches for a label text in the worksheet and writes to a cell relative to it.
     */
    function writeByLabel(ws, label, value, offsetCol = 1, offsetRow = 0, anchorFromEnd = true, maxCol = 16384) {
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
                    // If label is merged, we can anchor from start or end.
                    // General inputs want to jump past the label (anchor from end).
                    // Checkboxes want to tick the box to the left (anchor from start).
                    const merge = findMergeBounds(ws, cell.row, cell.col);
                    const baseCol = (anchorFromEnd && merge) ? merge.endCol : cell.col;

                    const targetRow = cell.row + offsetRow;
                    const targetCol = baseCol + offsetCol;

                    if (targetCol >= 1 && targetCol <= maxCol && targetRow >= 1) {
                        const targetCell = ws.getCell(targetRow, targetCol);
                        console.log(`Label [${label}] found at ${cell.address}. Writing "${value}" to ${targetCell.address}`);
                        targetCell.value = normalize(value);
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

    function fillRemedialInfo(sheet, startRow, remedial) {
        if (!sheet || !remedial) return;

        // Remedial header details
        if (remedial.from || remedial.to) {
            const period = `Conducted from ${remedial.from || ''} to ${remedial.to || ''}`;
            // Find where remedial header starts (usually row after final average)
            // But for now we use the startRow passed from caller
        }

        (remedial.subjects || []).forEach((subj, idx) => {
            const r = startRow + idx;
            // Col I (9): Subject
            sheet.getCell(r, 9).value = normalize(subj.subject);
            // Col BD (56): Sem Final
            sheet.getCell(r, 56).value = normalize(subj.semGrade);
            // Col BI (61): Remedial Mark
            sheet.getCell(r, 61).value = normalize(subj.remedialMark);
            // Col BN (66): Recomputed
            sheet.getCell(r, 66).value = normalize(subj.recomputedGrade);
            // Col BS (71): Action
            sheet.getCell(r, 71).value = normalize(subj.action);
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
            // Using smaller offsets now that findMergeBounds handles skipping the label's merged columns
            writeByLabel(front, 'DATE OF BIRTH', data.info?.birthdate, 1);
            writeByLabel(front, 'DATE OF SHS ADMISSION', data.info?.admissionDate, 1);

            // ELIGIBILITY Checkboxes (Offset -1 from START hits the box to the left)
            if (data.eligibility?.hsCompleter) writeByLabel(front, 'High School Completer*', 'X', -1, 0, false);
            if (data.eligibility?.jhsCompleter) writeByLabel(front, 'Junior High School Completer', 'X', -1, 0, false);
            if (data.eligibility?.pept) writeByLabel(front, 'PEPT Passer**', 'X', -1, 0, false);
            if (data.eligibility?.als) writeByLabel(front, 'ALS A&E Passer***', 'X', -1, 0, false);
            if (data.eligibility?.others) writeByLabel(front, 'Others (Pls. Specify):', 'X', -1, 0, false);

            writeByLabel(front, 'Gen. Ave', data.eligibility?.hsGenAve || data.eligibility?.jhsGenAve);
            writeByLabel(front, 'DATE OF GRADUATION', data.eligibility?.gradDate, 3);
            writeByLabel(front, 'NAME OF SCHOOL', data.eligibility?.schoolName, 2);
            writeByLabel(front, 'SCHOOL ADDRESS', data.eligibility?.schoolAddress, 1);

            // Additional Eligibility Details
            writeByLabel(front, 'Rating:', data.eligibility?.peptRating || data.eligibility?.alsRating);
            writeByLabel(front, 'Date of Examination/Assessment', data.eligibility?.examDate);
            writeByLabel(front, 'Name/Address of Community Learning Center', data.eligibility?.clcName);
            writeByLabel(front, 'Specify):', data.eligibility?.othersSpec);

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

            // Remedial Sections (Offsets from respective semester blocks)
            // Sem 1 Remedial: Row 55 on Front
            fillRemedialInfo(front, 55, data.semester1?.remedial);
            // Sem 2 Remedial: Row 98 on Front (Assuming 98 based on similar spacing)
            fillRemedialInfo(front, 98, data.semester2?.remedial);
            // Sem 3 Remedial: Row 37 on Back
            fillRemedialInfo(back, 37, data.semester3?.remedial);
            // Sem 4 Remedial: Row 78 on Back
            fillRemedialInfo(back, 78, data.semester4?.remedial);
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
            console.error('Bridge request error:', err);
            logToFile(`Request error: ${err.stack || err.message}`);
            return sendJson(req, res, 500, { success: false, error: err.message });
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
                return sendJson(req, res, 200, {
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

            return sendJson(req, res, 404, { success: false, error: 'Not found.' });
        } catch (err) {
            console.error('Server error:', err);
            return sendJson(req, res, 500, { success: false, error: err.message });
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
        }

        console.log(`Outputs:  ${OUTPUT_DIR}`);
        console.log('================================================');
        console.log('Keep this window open while using the website.');
        console.log('Press Ctrl+C to close.');
    });

} catch (err) {
    const errorMsg = `FATAL STARTUP ERROR: ${err.stack || err}\n`;
    console.error(errorMsg);
    logToFile(errorMsg);
    console.log('\nPress Enter to exit...');
    const rl = require('readline').createInterface({ input: process.stdin, output: process.stdout });
    rl.on('line', () => process.exit(1));
}
