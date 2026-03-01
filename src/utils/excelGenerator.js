import * as XLSX from 'xlsx';

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
                ws[target] = {
                    t: 's',
                    v: normalize(value)
                };
                return;
            }
        }
    }
}

function buildSummarySheet(data) {
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

    return XLSX.utils.aoa_to_sheet(rows);
}

function fillSemesterInfo(sheet, startRow, semData) {
    if (!sheet || !semData) return;
    const r = startRow - 1;
    // School
    sheet[XLSX.utils.encode_cell({ r, c: 4 })] = { t: 's', v: normalize(semData.school) };
    // School ID
    sheet[XLSX.utils.encode_cell({ r, c: 28 })] = { t: 's', v: normalize(semData.schoolId) };
    // Grade Level
    sheet[XLSX.utils.encode_cell({ r, c: 41 })] = { t: 's', v: normalize(semData.gradeLevel) };
    // SY
    sheet[XLSX.utils.encode_cell({ r, c: 52 })] = { t: 's', v: normalize(semData.sy) };
    // SEM
    sheet[XLSX.utils.encode_cell({ r, c: 62 })] = { t: 's', v: normalize(semData.semester) };
    // Track/Strand
    sheet[XLSX.utils.encode_cell({ r: r + 2, c: 6 })] = { t: 's', v: normalize(semData.trackStrand) };
    // Section
    sheet[XLSX.utils.encode_cell({ r: r + 2, c: 42 })] = { t: 's', v: normalize(semData.section) };
}

function fillSemesterSubjects(sheet, startRow, subjects) {
    if (!sheet || !subjects || !Array.isArray(subjects)) return;
    subjects.forEach((subj, idx) => {
        const r = startRow - 1 + idx;
        sheet[XLSX.utils.encode_cell({ r, c: 0 })] = { t: 's', v: normalize(subj.type) };
        sheet[XLSX.utils.encode_cell({ r, c: 8 })] = { t: 's', v: normalize(subj.subject) };
        sheet[XLSX.utils.encode_cell({ r, c: 45 })] = { t: 's', v: normalize(subj.q1) };
        sheet[XLSX.utils.encode_cell({ r, c: 50 })] = { t: 's', v: normalize(subj.q2) };
        sheet[XLSX.utils.encode_cell({ r, c: 55 })] = { t: 's', v: normalize(subj.final) };
        sheet[XLSX.utils.encode_cell({ r, c: 60 })] = { t: 's', v: normalize(subj.action) };
    });
}

export async function generateExcelForm(data) {
    try {
        const response = await fetch('/Form 137-SHS-BLANK.xlsx');

        let workbook = null;
        if (response.ok) {
            const arrayBuffer = await response.arrayBuffer();
            workbook = XLSX.read(arrayBuffer, {
                type: 'array',
                cellStyles: true,
                cellFormula: true
            });

            const sheetNames = workbook.SheetNames || [];
            const frontName = sheetNames.find(name => name.toUpperCase().includes('FRONT')) || sheetNames[0];
            const backName = sheetNames.find(name => name.toUpperCase().includes('BACK')) || sheetNames[1] || frontName;

            const front = frontName ? workbook.Sheets[frontName] : null;
            const back = backName ? workbook.Sheets[backName] : null;

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

                fillSemesterInfo(front, 23, data.semester1);
                fillSemesterSubjects(front, 28, data.semester1?.subjects);
                fillSemesterInfo(front, 66, data.semester2);
                fillSemesterSubjects(front, 71, data.semester2?.subjects);
            }

            if (back) {
                writeByLabel(back, 'TRACK/STRAND', data.certification?.trackStrand, 3);
                writeByLabel(back, 'SHS GENERAL AVERAGE', data.certification?.genAve, 3);
                writeByLabel(back, 'DATE OF GRADUATION', data.certification?.gradDate, 2, 2);
                writeByLabel(back, 'NAME OF SCHOOL', data.certification?.schoolHead, 2);

                fillSemesterInfo(back, 4, data.semester3);
                fillSemesterSubjects(back, 11, data.semester3?.subjects);
                fillSemesterInfo(back, 46, data.semester4);
                fillSemesterSubjects(back, 51, data.semester4?.subjects);
            }

            const filename = `Form137_${(data.info?.lname || 'Student').replace(/[^a-z0-9]/gi, '_')}.xlsx`;
            XLSX.writeFile(workbook, filename);
            return { success: true };
        }

        workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, buildSummarySheet(data), 'Form137');
        XLSX.writeFile(workbook, `Form137_${Date.now()}.xlsx`);
        return { success: true };
    } catch (error) {
        console.error('Error generating Excel file:', error);
        return { success: false, error: error.message };
    }
}
