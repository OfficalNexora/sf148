import ExcelJS from 'exceljs';

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

    // ExcelJS iterate rows
    ws.eachRow((row) => {
        if (found) return;
        row.eachCell((cell) => {
            if (found) return;
            const text = normalize(cell.value).toUpperCase();
            if (text.includes(search)) {
                const targetRow = cell.row + offsetRow;
                const targetCol = cell.col + offsetCol;
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
    const r = startRow;
    sheet.getCell(r, 5).value = normalize(semData.school);
    sheet.getCell(r, 29).value = normalize(semData.schoolId);
    sheet.getCell(r, 42).value = normalize(semData.gradeLevel);
    sheet.getCell(r, 53).value = normalize(semData.sy);
    sheet.getCell(r, 63).value = normalize(semData.semester);
    sheet.getCell(r + 2, 7).value = normalize(semData.trackStrand);
    sheet.getCell(r + 2, 43).value = normalize(semData.section);
}

function fillSemesterSubjects(sheet, startRow, subjects) {
    if (!sheet || !subjects || !Array.isArray(subjects)) return;
    subjects.forEach((subj, idx) => {
        const r = startRow + idx;
        if (idx > 100) return; // Safety
        sheet.getCell(r, 1).value = normalize(subj.type);
        sheet.getCell(r, 9).value = normalize(subj.subject);
        sheet.getCell(r, 46).value = normalize(subj.q1);
        sheet.getCell(r, 51).value = normalize(subj.q2);
        sheet.getCell(r, 56).value = normalize(subj.final);
        sheet.getCell(r, 61).value = normalize(subj.action);
    });
}

function buildSummarySheet(workbook, data) {
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
}

async function triggerDownload(workbook, filename) {
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    window.URL.revokeObjectURL(url);
}

export async function generateExcelForm(data) {
    try {
        const response = await fetch('/Form 137-SHS-BLANK.xlsx');
        const filename = `Form137_${(data.info?.lname || 'Student').replace(/[^a-z0-9]/gi, '_')}.xlsx`;

        if (response.ok) {
            const arrayBuffer = await response.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);

            const sheets = workbook.worksheets;
            const front = sheets.find(s => s.name.toUpperCase().includes('FRONT')) || sheets[0];
            const back = sheets.find(s => s.name.toUpperCase().includes('BACK')) || sheets[1] || front;

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

            await triggerDownload(workbook, filename);
            return { success: true };
        }

        const workbook = new ExcelJS.Workbook();
        buildSummarySheet(workbook, data);
        await triggerDownload(workbook, `Form137_Summary_${Date.now()}.xlsx`);
        return { success: true };
    } catch (error) {
        console.error('Error generating Excel file:', error);
        return { success: false, error: error.message };
    }
}
