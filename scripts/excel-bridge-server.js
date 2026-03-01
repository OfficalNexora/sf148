/* eslint-disable no-console */
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { exec, execFile } = require('child_process');
const XLSX = require('xlsx');

const PORT = Number(process.env.BRIDGE_PORT || 8787);
const HOST = process.env.BRIDGE_HOST || '127.0.0.1';
const API_KEY = process.env.BRIDGE_API_KEY || '';
const ALLOW_ORIGIN = process.env.BRIDGE_ALLOW_ORIGIN || '*';
const TEMPLATE_PATH = process.env.BRIDGE_TEMPLATE_PATH || path.join(process.cwd(), 'public', 'Form 137-SHS-BLANK.xlsx');
const OUTPUT_DIR = process.env.BRIDGE_OUTPUT_DIR || path.join(os.tmpdir(), 'form137-exports');

if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
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

function writeByLabel(ws, label, value, offsetCol = 1, offsetRow = 0, maxCol = 24) {
    if (value === undefined || value === null || value === '') return;
    if (!ws || !ws['!ref']) return;

    const search = label.toUpperCase();
    const range = XLSX.utils.decode_range(ws['!ref']);
    const endCol = Math.min(range.e.c, maxCol);

    for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= endCol; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            const text = (cell && cell.v !== undefined && cell.v !== null)
                ? String(cell.v).toUpperCase()
                : '';

            if (text.includes(search)) {
                const target = XLSX.utils.encode_cell({ r: r + offsetRow, c: c + offsetCol });
                ws[target] = { t: 's', v: normalize(value) };
                return;
            }
        }
    }
}

function createSummaryWorkbook(data) {
    const wb = XLSX.utils.book_new();
    const rows = [];
    rows.push(['FORM 137 - Student Permanent Record']);
    rows.push([]);
    rows.push(['Last Name', data.info?.lname || '']);
    rows.push(['First Name', data.info?.fname || '']);
    rows.push(['Middle Name', data.info?.mname || '']);
    rows.push(['LRN', data.info?.lrn || '']);
    rows.push(['Sex', data.info?.sex || '']);
    rows.push(['Date of Birth', data.info?.birthdate || '']);
    rows.push([]);

    const semesterKeys = ['semester1', 'semester2', 'semester3', 'semester4'];
    semesterKeys.forEach((semKey, index) => {
        const sem = data[semKey];
        if (!sem) return;

        rows.push([`Semester ${index + 1}`]);
        rows.push(['School', sem.school || '']);
        rows.push(['School Year', sem.sy || '']);
        rows.push(['Grade Level', sem.gradeLevel || '']);
        rows.push(['Track/Strand', sem.trackStrand || '']);
        rows.push(['Section', sem.section || '']);
        rows.push(['Type', 'Subject', 'Q1', 'Q2', 'Final', 'Action']);

        (sem.subjects || []).forEach((subj) => {
            if (!subj || !subj.subject) return;
            rows.push([
                subj.type || '',
                subj.subject || '',
                subj.q1 || '',
                subj.q2 || '',
                subj.final || '',
                subj.action || ''
            ]);
        });

        rows.push(['General Average', sem.genAve || '']);
        rows.push(['Remarks', sem.remarks || '']);
        rows.push([]);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Form137');
    return wb;
}

function createWorkbookFromTemplate(data) {
    if (!fs.existsSync(TEMPLATE_PATH)) {
        return createSummaryWorkbook(data);
    }

    const wb = XLSX.readFile(TEMPLATE_PATH, {
        cellStyles: true,
        cellFormula: true
    });

    const sheetNames = wb.SheetNames || [];
    const frontName = sheetNames.find(name => name.toUpperCase().includes('FRONT')) || sheetNames[0];
    const backName = sheetNames.find(name => name.toUpperCase().includes('BACK')) || sheetNames[1] || frontName;

    const front = frontName ? wb.Sheets[frontName] : null;
    const back = backName ? wb.Sheets[backName] : null;

    if (front) {
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
    }

    if (back) {
        writeByLabel(back, 'TRACK/STRAND', data.certification?.trackStrand, 3);
        writeByLabel(back, 'SHS GENERAL AVERAGE', data.certification?.genAve, 3);
        writeByLabel(back, 'DATE OF GRADUATION', data.certification?.gradDate, 2, 2);
        writeByLabel(back, 'NAME OF SCHOOL', data.certification?.schoolHead, 2);
    }

    return wb;
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
        const psScript = [
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

        execFile(
            'powershell.exe',
            ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', psScript],
            (error, stdout, stderr) => {
                if (error) {
                    const details = (stderr || stdout || error.message || '').trim();
                    reject(new Error(details || 'PowerShell print failed.'));
                    return;
                }
                resolve();
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

    const body = await parseBody(req);
    const data = body && body.data;
    if (!data || !data.info) {
        return sendJson(res, 400, { success: false, error: 'Missing student data.' });
    }

    const autoPrint = Boolean(body && body.autoPrint);
    const openAfterPrint = body && body.openAfterPrint !== false;

    const workbook = createWorkbookFromTemplate(data);
    const lastName = (data.info?.lname || 'Student').replace(/[^a-z0-9]/gi, '_');
    const filePath = path.join(OUTPUT_DIR, `Form137_${lastName}_${Date.now()}.xlsx`);
    XLSX.writeFile(workbook, filePath);

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
        console.error('Bridge error:', err);
        return sendJson(res, 500, { success: false, error: err.message });
    }
});

server.listen(PORT, HOST, () => {
    console.log(`Excel bridge listening on http://${HOST}:${PORT}`);
    console.log(`Template path: ${TEMPLATE_PATH}`);
    console.log(`Output dir: ${OUTPUT_DIR}`);
    if (API_KEY) console.log('API key auth: enabled');
});
