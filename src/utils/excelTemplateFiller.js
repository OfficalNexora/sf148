import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// Helper to convert grades to numbers
const toNum = (val) => {
    if (val === '' || val === null || val === undefined) return '';
    const n = parseFloat(val);
    return isNaN(n) ? val : n;
};

// Map grade types to template columns
const TYPE_MAP = {
    'Core': 'A',
    'Applied': 'J',
    'Specialized': 'S',
    'Other': 'A' // Default fallback
};

export const generateExcelFromTemplate = async (data) => {
    try {
        // 1. Fetch the template from the public folder
        const response = await fetch('/PLACEHOLDER(ALL).xlsx');
        if (!response.ok) throw new Error('Could not find PLACEHOLDER(ALL).xlsx template in public folder.');

        const arrayBuffer = await response.arrayBuffer();

        // 2. Load into ExcelJS
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        const wsFront = workbook.getWorksheet('FRONT') || workbook.worksheets[0];
        const wsBack = workbook.getWorksheet('BACK') || (workbook.worksheets.length > 1 ? workbook.worksheets[1] : workbook.worksheets[0]);

        // Helper to safely write to a cell
        const safeWrite = (ws, cellAddress, value) => {
            if (value !== undefined && value !== null && value !== '') {
                ws.getCell(cellAddress).value = value;
            }
        };

        // Helper for subject rows
        const writeSubjectRow = (ws, rowIdx, subj) => {
            if (!subj || !subj.subject) return;
            const t = subj.type || 'Core';
            let col = TYPE_MAP[t] || 'A';
            safeWrite(ws, `${col}${rowIdx}`, subj.subject);
            safeWrite(ws, `AR${rowIdx}`, toNum(subj.q1));
            safeWrite(ws, `AY${rowIdx}`, toNum(subj.q2));
            safeWrite(ws, `BD${rowIdx}`, toNum(subj.final));
            safeWrite(ws, `BJ${rowIdx}`, subj.action);
        };

        // Helper for remedial rows
        const writeRemedialRow = (ws, rowIdx, rem) => {
            if (!rem || !rem.subject) return;
            safeWrite(ws, `A${rowIdx}`, rem.subject);
            safeWrite(ws, `X${rowIdx}`, toNum(rem.semGrade));
            safeWrite(ws, `AH${rowIdx}`, toNum(rem.remedialMark));
            safeWrite(ws, `AR${rowIdx}`, toNum(rem.recomputedGrade));
            safeWrite(ws, `BB${rowIdx}`, rem.action);
        };

        const { info = {}, eligibility = {}, certification = {} } = data;

        // --- STUDENT INFO (FRONT) ---
        safeWrite(wsFront, 'F8', info.lname);
        safeWrite(wsFront, 'Y8', info.fname);
        safeWrite(wsFront, 'AZ8', info.mname);
        safeWrite(wsFront, 'C9', info.lrn);
        safeWrite(wsFront, 'AN9', info.sex);
        safeWrite(wsFront, 'AA9', info.birthdate);
        safeWrite(wsFront, 'BH9', info.admissionDate);

        // --- ELIGIBILITY (FRONT) ---
        safeWrite(wsFront, 'Z14', eligibility.schoolName);
        safeWrite(wsFront, 'AW14', eligibility.schoolAddress);
        if (eligibility.hsCompleter) safeWrite(wsFront, 'A13', 'X');
        safeWrite(wsFront, 'N13', eligibility.hsGenAve);
        if (eligibility.jhsCompleter) safeWrite(wsFront, 'S13', 'X');
        safeWrite(wsFront, 'AH13', eligibility.jhsGenAve);
        safeWrite(wsFront, 'P14', eligibility.gradDate);
        if (eligibility.pept) safeWrite(wsFront, 'A16', 'X');
        safeWrite(wsFront, 'K16', eligibility.peptRating);
        if (eligibility.als) safeWrite(wsFront, 'S16', 'X');
        safeWrite(wsFront, 'AC16', eligibility.alsRating);
        if (eligibility.others) safeWrite(wsFront, 'AH16', 'X');
        safeWrite(wsFront, 'AP16', eligibility.othersSpec);
        safeWrite(wsFront, 'P17', eligibility.examDate);
        safeWrite(wsFront, 'AN17', eligibility.clcName);

        // --- SEMESTER 1 ---
        const s1 = data.semester1 || {};
        safeWrite(wsFront, 'E23', s1.school); safeWrite(wsFront, 'AF23', s1.schoolId);
        safeWrite(wsFront, 'AS23', s1.gradeLevel); safeWrite(wsFront, 'BA23', s1.sy);
        safeWrite(wsFront, 'BK23', s1.sem); safeWrite(wsFront, 'G25', s1.trackStrand);
        safeWrite(wsFront, 'AS25', s1.section);
        const sub1 = s1.subjects || [];
        for (let i = 0; i < 11; i++) writeSubjectRow(wsFront, 31 + i, sub1[i]);
        safeWrite(wsFront, 'BD43', toNum(s1.genAve));
        safeWrite(wsFront, 'E45', s1.remarks);
        safeWrite(wsFront, 'A49', s1.adviserName);
        safeWrite(wsFront, 'Y49', s1.certName);
        safeWrite(wsFront, 'AZ49', s1.dateChecked);

        const r1 = s1.remedial || {};
        if (r1.subjects?.length) {
            safeWrite(wsFront, 'S52', r1.from); safeWrite(wsFront, 'AC52', r1.to);
            safeWrite(wsFront, 'AL52', r1.school); safeWrite(wsFront, 'BK52', r1.schoolId);
            for (let i = 0; i < 5; i++) writeRemedialRow(wsFront, 58 + i, r1.subjects[i]);
            safeWrite(wsFront, 'J63', r1.teacherName);
            safeWrite(wsFront, 'AY63', r1.signature); // Not typical in current json but keeping logic
        }

        // --- SEMESTER 2 ---
        const s2 = data.semester2 || {};
        safeWrite(wsFront, 'E66', s2.school); safeWrite(wsFront, 'AF66', s2.schoolId);
        safeWrite(wsFront, 'AS66', s2.gradeLevel); safeWrite(wsFront, 'BA66', s2.sy);
        safeWrite(wsFront, 'BK66', s2.sem); safeWrite(wsFront, 'G68', s2.trackStrand);
        safeWrite(wsFront, 'AS68', s2.section);
        const sub2 = s2.subjects || [];
        for (let i = 0; i < 11; i++) writeSubjectRow(wsFront, 74 + i, sub2[i]);
        safeWrite(wsFront, 'BD86', toNum(s2.genAve));
        safeWrite(wsFront, 'F88', s2.remarks);
        safeWrite(wsFront, 'A92', s2.adviserName);
        safeWrite(wsFront, 'Y92', s2.certName);
        safeWrite(wsFront, 'AZ92', s2.dateChecked);

        const r2 = s2.remedial || {};
        if (r2.subjects?.length) {
            safeWrite(wsFront, 'S95', r2.from); safeWrite(wsFront, 'AC95', r2.to);
            safeWrite(wsFront, 'AL95', r2.school); safeWrite(wsFront, 'BK95', r2.schoolId);
            for (let i = 0; i < 5; i++) writeRemedialRow(wsFront, 101 + i, r2.subjects[i]);
            safeWrite(wsFront, 'J106', r2.teacherName);
            safeWrite(wsFront, 'AY106', r2.signature);
        }

        // --- SEMESTER 3 (BACK) ---
        const s3 = data.semester3 || {};
        safeWrite(wsBack, 'E4', s3.school); safeWrite(wsBack, 'AF4', s3.schoolId);
        safeWrite(wsBack, 'AS4', s3.gradeLevel); safeWrite(wsBack, 'BA4', s3.sy);
        safeWrite(wsBack, 'BK4', s3.sem); safeWrite(wsBack, 'G5', s3.trackStrand);
        safeWrite(wsBack, 'AS5', s3.section);
        const sub3 = s3.subjects || [];
        for (let i = 0; i < 12; i++) writeSubjectRow(wsBack, 11 + i, sub3[i]);
        safeWrite(wsBack, 'BD23', toNum(s3.genAve));
        safeWrite(wsBack, 'F25', s3.remarks);
        safeWrite(wsBack, 'A29', s3.adviserName);
        safeWrite(wsBack, 'Y29', s3.certName);
        safeWrite(wsBack, 'AZ29', s3.dateChecked);

        const r3 = s3.remedial || {};
        if (r3.subjects?.length) {
            safeWrite(wsBack, 'S32', r3.from); safeWrite(wsBack, 'AC32', r3.to);
            safeWrite(wsBack, 'AL32', r3.school); safeWrite(wsBack, 'BK32', r3.schoolId);
            for (let i = 0; i < 4; i++) writeRemedialRow(wsBack, 38 + i, r3.subjects[i]);
            safeWrite(wsBack, 'J43', r3.teacherName);
            safeWrite(wsBack, 'AY43', r3.signature);
        }

        // --- SEMESTER 4 (BACK) ---
        const s4 = data.semester4 || {};
        safeWrite(wsBack, 'E46', s4.school); safeWrite(wsBack, 'AF46', s4.schoolId);
        safeWrite(wsBack, 'AS46', s4.gradeLevel); safeWrite(wsBack, 'BA46', s4.sy);
        safeWrite(wsBack, 'BK46', s4.sem); safeWrite(wsBack, 'G48', s4.trackStrand);
        safeWrite(wsBack, 'BC48', s4.section);
        const sub4 = s4.subjects || [];
        for (let i = 0; i < 12; i++) writeSubjectRow(wsBack, 54 + i, sub4[i]);
        safeWrite(wsBack, 'BD66', toNum(s4.genAve));
        safeWrite(wsBack, 'F68', s4.remarks);
        safeWrite(wsBack, 'A72', s4.adviserName);
        safeWrite(wsBack, 'Y72', s4.certName);
        safeWrite(wsBack, 'AZ72', s4.dateChecked);

        const r4 = s4.remedial || {};
        if (r4.subjects?.length) {
            safeWrite(wsBack, 'S75', r4.from); safeWrite(wsBack, 'AC75', r4.to);
            safeWrite(wsBack, 'AL75', r4.school); safeWrite(wsBack, 'BK75', r4.schoolId);
            for (let i = 0; i < 4; i++) writeRemedialRow(wsBack, 81 + i, r4.subjects[i]);
            safeWrite(wsBack, 'I86', r4.teacherName);
            safeWrite(wsBack, 'AY86', r4.signature);
        }

        // --- CERTIFICATION ---
        safeWrite(wsBack, 'BI91', certification.gradDate);
        safeWrite(wsBack, 'BJ90', certification.genAve);
        safeWrite(wsBack, 'A94', certification.schoolHead);
        safeWrite(wsBack, 'I90', certification.trackStrand);
        safeWrite(wsBack, 'I91', certification.awards);
        safeWrite(wsBack, 'A112', certification.remarks);
        safeWrite(wsBack, 'T94', certification.certDate);
        safeWrite(wsBack, 'J114', certification.dateIssued);

        // 3. Export to File
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const safeName = (info.lname || 'Student').replace(/\s+/g, '_');
        const filename = `SF10_${safeName}_${Date.now()}.xlsx`;

        saveAs(blob, filename);

        return { success: true, filename };

    } catch (error) {
        console.error("Excel Generation Error:", error);
        return { success: false, error: error.message };
    }
};
