// excel-generator-worker.js
// Uses exceljs to generate/parse files using placeholders like %(lname)

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Mapping of placeholders to data object paths
const PLACEHOLDER_MAP = {
    'lname': (d) => d.info.lname,
    'fname': (d) => d.info.fname,
    'mname': (d) => d.info.mname,
    'name': (d) => `${d.info.lname}, ${d.info.fname}`,
    'full_name': (d) => `${d.info.fname} ${d.info.mname ? d.info.mname + ' ' : ''}${d.info.lname}`,
    'lrn': (d) => d.info.lrn,
    'bdate': (d) => d.info.birthdate,
    'sex': (d) => d.info.sex,
    'admission_date': (d) => d.info.admissionDate,
    'school_name': (d) => d.eligibility.schoolName,
    'school_addr': (d) => d.eligibility.schoolAddress,
    'hs_ave': (d) => d.eligibility.hsGenAve,
    'jhs_ave': (d) => d.eligibility.jhsGenAve,
    'grad_date': (d) => d.eligibility.gradDate,
    'grade_level': (d) => d.semester1.gradeLevel,
    'section': (d) => d.semester1.section,
    'track_strand': (d) => d.semester1.trackStrand,
};

// Reverse mapping for parsing: (placeholder) => (setter function)
const PARSE_SETTERS = {
    'lname': (d, v) => d.info.lname = v,
    'fname': (d, v) => d.info.fname = v,
    'mname': (d, v) => d.info.mname = v,
    'lrn': (d, v) => d.info.lrn = v,
    'bdate': (d, v) => d.info.birthdate = v,
    'sex': (d, v) => d.info.sex = v,
    'school_name': (d, v) => d.eligibility.schoolName = v,
    'school_addr': (d, v) => d.eligibility.schoolAddress = v,
    'grade_level': (d, v) => { d.semester1.gradeLevel = v; d.semester2.gradeLevel = v; },
    'section': (d, v) => { d.semester1.section = v; d.semester2.section = v; },
    'track_strand': (d, v) => { d.semester1.trackStrand = v; d.semester2.trackStrand = v; },
};

process.on('message', async (msg) => {
    if (msg.type === 'generate-print-file') {
        const { data, templatePath } = msg;

        try {
            if (!fs.existsSync(templatePath)) throw new Error('Template not found');

            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(templatePath);

            // Iterate through ALL cells in the first worksheet to find and replace placeholders
            const ws = wb.worksheets[0];
            if (!ws) throw new Error('No worksheets found');

            ws.eachRow((row) => {
                row.eachCell((cell) => {
                    const val = cell.value;
                    if (val && typeof val === 'string' && val.includes('%(')) {
                        // Regex to find multiple placeholders in one cell if needed
                        const regex = /%\((.*?)\)/g;
                        let match;
                        let cellStr = val;
                        let replaced = false;

                        while ((match = regex.exec(val)) !== null) {
                            const key = match[1];
                            const getter = PLACEHOLDER_MAP[key];
                            if (getter) {
                                const dataVal = getter(data) || '';
                                cellStr = cellStr.replace(`%(${key})`, dataVal);
                                replaced = true;
                            }
                        }
                        if (replaced) {
                            cell.value = cellStr;
                        }
                    }
                });
            });

            const tempDir = require('os').tmpdir();
            const studentName = (data.info.lname || 'Form137').replace(/[^a-z0-9]/gi, '_');
            const outputPath = path.join(tempDir, `${studentName}_${Date.now()}.xlsx`);

            await wb.xlsx.writeFile(outputPath);
            process.send({ type: 'result', success: true, filePath: outputPath });

        } catch (e) {
            process.send({ type: 'result', success: false, error: e.message });
        }
    }

    if (msg.type === 'parse-excel') {
        const { filePath } = msg;

        try {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(filePath);
            const ws = wb.worksheets[0];
            if (!ws) throw new Error('No worksheets found');

            // Start with a generic empty structure
            // We'll need to improve this to avoid losing data during sync
            const data = {
                info: { lname: '', fname: '', mname: '', lrn: '', sex: '', birthdate: '', admissionDate: '' },
                eligibility: { schoolName: '', schoolAddress: '' },
                semester1: { subjects: [] },
                semester2: { subjects: [] },
                semester3: { subjects: [] },
                semester4: { subjects: [] },
                certification: {}
            };

            // SCAN for placeholders and grab values from those cells
            // Wait, if it's a SYNC/IMPORT, the cell might ALREADY contain value
            // but how do we know which cell was which placeholder if the placeholder text is GONE?

            // ARCHITECTURAL NOTE: For placeholders to work for PARSING, 
            // the template must be read twice (one blank, one filled) 
            // OR we use the previous label-based logic as a fallback.
            // BETTER: If the user provides a "filled" form, we assume the structure is FIXED.

            // For now, let's stick to label-based fallback for parsing 
            // UNLESS the placeholder text is somehow preserved (unlikely in a final form).

            // HEURISTIC: We'll use the Label-based scan but also check if we can improve it.
            const findVal = (label, offsetCol = 1) => {
                const searchLabel = label.toUpperCase();
                for (let r = 1; r <= 80; r++) {
                    const row = ws.getRow(r);
                    for (let c = 1; c <= 26; c++) {
                        const cell = row.getCell(c);
                        if (cell.value && cell.value.toString().toUpperCase().includes(searchLabel)) {
                            let targetCell = row.getCell(c + offsetCol);
                            if (targetCell.isMerged && targetCell.master) targetCell = targetCell.master;
                            const v = targetCell.value;
                            if (v && typeof v === 'object' && v.richText) return v.richText.map(t => t.text).join('');
                            return v ? v.toString().trim() : '';
                        }
                    }
                }
                return '';
            };

            data.info.lname = findVal('LAST NAME', 1);
            data.info.fname = findVal('FIRST NAME', 1);
            data.info.mname = findVal('MIDDLE NAME', 1);
            data.info.lrn = findVal('LRN', 1);
            data.info.birthdate = findVal('DATE OF BIRTH', 3);
            data.info.sex = findVal('SEX', 1);
            data.eligibility.schoolName = findVal('Name of School', 1);
            data.eligibility.schoolAddress = findVal('School Address', 1);

            process.send({ type: 'parse-result', success: true, data });

        } catch (e) {
            process.send({ type: 'parse-result', success: false, error: e.message });
        }
    } else if (msg.type === 'scan-placeholders') {
        const { templatePath } = msg;
        try {
            if (!fs.existsSync(templatePath)) {
                throw new Error('Template not found');
            }
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(templatePath);

            const placeholders = new Set();
            wb.eachSheet((ws) => {
                ws.eachRow((row) => {
                    row.eachCell((cell) => {
                        const val = cell.value;
                        if (typeof val === 'string' && val.includes('%(')) {
                            // Extract all occurrences like %(field)
                            const matches = val.match(/%\(([a-zA-Z0-9_]+)\)/g);
                            if (matches) {
                                matches.forEach(m => placeholders.add(m.slice(2, -1)));
                            }
                        }
                    });
                });
            });

            process.send({ type: 'scan-result', success: true, placeholders: Array.from(placeholders) });
        } catch (e) {
            process.send({ type: 'scan-result', success: false, error: e.message });
        }
    }
});
