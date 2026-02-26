import ExcelJS from 'exceljs';

export async function generateExcelForm(data) {
    try {
        // Fetch template file from public folder
        const response = await fetch('/Form 137-SHS-BLANK.xlsx');
        if (!response.ok) throw new Error('Could not find the Excel template.');

        const arrayBuffer = await response.arrayBuffer();

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(arrayBuffer);

        const front = wb.getWorksheet('FRONT');
        const back = wb.getWorksheet('BACK');
        const annexSheet = wb.getWorksheet('ANNEX');

        if (!front || !back) {
            throw new Error('FRONT or BACK worksheet missing in the template.');
        }

        // --- HELPER: Find label and write relative to it (Fallback) ---
        const writeToLabel = (ws, label, value, offsetCol = 1, offsetRow = 0) => {
            if (value === undefined || value === null || value === '') return;
            const search = label.toUpperCase();
            ws.eachRow((row, rowNum) => {
                row.eachCell((cell, colNum) => {
                    const cellVal = cell.value ? cell.value.toString().toUpperCase() : '';
                    if (cellVal.includes(search)) {
                        // Only write if target is EMPTY
                        const targetRow = ws.getRow(rowNum + offsetRow);
                        const targetCell = targetRow.getCell(colNum + offsetCol);
                        if (!targetCell.value) targetCell.value = value;
                    }
                });
            });
        };

        // --- HELPER: Find subject data across semesters (Dynamic Retrieval) ---
        const findSubjectData = (subjectName) => {
            if (!subjectName || subjectName.trim() === '') return null;
            const search = subjectName.trim().toUpperCase();

            for (let i = 1; i <= 4; i++) {
                const sem = data[`semester${i}`];
                if (sem && sem.subjects) {
                    const match = sem.subjects.find(s => s.subject && s.subject.trim().toUpperCase() === search);
                    if (match) return { q1: match.q1, q2: match.q2, final: match.final, action: match.action };
                }
                if (sem && sem.remedial && sem.remedial.subjects) {
                    const match = sem.remedial.subjects.find(s => s.subject && s.subject.trim().toUpperCase() === search);
                    if (match) return { q1: match.semGrade, q2: match.remedialMark, final: match.recomputedGrade, action: match.action };
                }
            }
            return null;
        };

        // --- HELPER: Direct Placeholder Replacement ---
        const placeholderMap = {
            'lname': data.info.lname,
            'fname': data.info.fname,
            'mname': data.info.mname,
            'lrn': data.info.lrn,
            'sex': data.info.sex,
            'birthdate': data.info.birthdate,
            'admission_date': data.info.admissionDate,
            'hs_school': data.eligibility.schoolName,
            'hs_addr': data.eligibility.schoolAddress,
            'hs_ave': data.eligibility.hsGenAve,
            'grad_date': data.eligibility.gradDate,
            // Sem 1
            's1_school': data.semester1.school,
            's1_id': data.semester1.schoolId,
            's1_level': data.semester1.gradeLevel,
            's1_sy': data.semester1.sy,
            's1_sem': data.semester1.sem,
            's1_ave': data.semester1.genAve,
            // Compatibility Aliases 
            'lastname': data.info.lname,
            'firstname': data.info.fname,
            'middlename': data.info.mname,
            'LRN': data.info.lrn,
            'date of birth': data.info.birthdate,
            'dateofadmission': data.info.admissionDate,
            'genave': data.semester1.genAve,
            'track': data.semester1.trackStrand,
            'strand': data.semester1.trackStrand,
            'section': data.semester1.section,
            // Sem 2
            's2_school': data.semester2.school,
            's2_id': data.semester2.schoolId,
            's2_level': data.semester2.gradeLevel,
            's2_sy': data.semester2.sy,
            's2_sem': data.semester2.sem,
            's2_ave': data.semester2.genAve,
            // Sem 3
            's3_school': data.semester3.school,
            's3_id': data.semester3.schoolId,
            's3_level': data.semester3.gradeLevel,
            's3_sy': data.semester3.sy,
            's3_sem': data.semester3.sem,
            's3_ave': data.semester3.genAve,
            // Sem 4
            's4_school': data.semester4.school,
            's4_id': data.semester4.schoolId,
            's4_level': data.semester4.gradeLevel,
            's4_sy': data.semester4.sy,
            's4_sem': data.semester4.sem,
            's4_ave': data.semester4.genAve,
        };

        // Dynamic Subject Placeholders for Semester 1-4 (Max 10 subjects each)
        [1, 2, 3, 4].forEach(sNum => {
            const semData = data[`semester${sNum}`];
            if (semData && semData.subjects) {
                semData.subjects.forEach((subj, i) => {
                    const idx = i + 1;
                    placeholderMap[`s${sNum}sub_${idx}`] = subj.subject;
                    placeholderMap[`s${sNum}q1_${idx}`] = subj.q1;
                    placeholderMap[`s${sNum}q2_${idx}`] = subj.q2;
                    placeholderMap[`s${sNum}fin_${idx}`] = subj.final;
                    placeholderMap[`s${sNum}act_${idx}`] = subj.action;
                });
                if (semData.remedial && semData.remedial.subjects) {
                    semData.remedial.subjects.forEach((subj, i) => {
                        const idx = i + 1;
                        placeholderMap[`s${sNum}rem_sub_${idx}`] = subj.subject;
                        placeholderMap[`s${sNum}rem_q1_${idx}`] = subj.semGrade;
                        placeholderMap[`s${sNum}rem_q2_${idx}`] = subj.remedialMark;
                        placeholderMap[`s${sNum}rem_fin_${idx}`] = subj.recomputedGrade;
                        placeholderMap[`s${sNum}rem_act_${idx}`] = subj.action;
                    });
                }
            }
        });

        // Dynamic Annex Placeholders
        if (data.annex) {
            data.annex.forEach((subj, i) => {
                const idx = i + 1;
                const retrieved = findSubjectData(subj.subject);

                placeholderMap[`asub_${idx}`] = subj.subject;
                placeholderMap[`atype_${idx}`] = subj.type;
                placeholderMap[`aq1_${idx}`] = retrieved ? retrieved.q1 : '';
                placeholderMap[`aq2_${idx}`] = retrieved ? retrieved.q2 : '';
                placeholderMap[`afin_${idx}`] = retrieved ? retrieved.final : '';
                placeholderMap[`aact_${idx}`] = retrieved ? retrieved.action : '';
            });
        }

        const replacePlaceholders = (ws) => {
            ws.eachRow((row) => {
                row.eachCell((cell) => {
                    let val = cell.value ? cell.value.toString() : '';
                    if (val.includes('%(')) {
                        Object.keys(placeholderMap).forEach(key => {
                            const p = `%(${key})`;
                            if (val.includes(p)) {
                                val = val.replace(p, placeholderMap[key] || '');
                            }
                        });
                        cell.value = val;
                    }
                });
            });
        };

        if (front) replacePlaceholders(front);
        if (back) replacePlaceholders(back);
        if (annexSheet) replacePlaceholders(annexSheet);

        // --- FILL FRONT SHEET (Fallback) ---
        writeToLabel(front, 'LAST NAME:', data.info.lname);
        writeToLabel(front, 'FIRST NAME:', data.info.fname);
        writeToLabel(front, 'MIDDLE NAME:', data.info.mname);
        writeToLabel(front, 'LRN:', data.info.lrn);
        writeToLabel(front, 'SEX:', data.info.sex);
        writeToLabel(front, 'DATE OF BIRTH', data.info.birthdate, 3);
        writeToLabel(front, 'DATE OF SHS ADMISSION', data.info.admissionDate, 5);

        if (data.eligibility.hsCompleter) writeToLabel(front, 'High School Completer', 'X', -1);
        if (data.eligibility.jhsCompleter) writeToLabel(front, 'Junior High School Completer', 'X', -1);
        writeToLabel(front, 'Gen. Ave:', data.eligibility.hsGenAve);
        writeToLabel(front, 'Date of Graduation', data.eligibility.gradDate, 3);
        writeToLabel(front, 'Name of School:', data.eligibility.schoolName, 2);
        writeToLabel(front, 'School Address:', data.eligibility.schoolAddress, 1);

        // --- FILL SEMESTERS (Fallback Coordinates) ---
        const fillSemester = (ws, semData, startLabelSearch) => {
            if (!semData) return;
            let startRow = -1;
            let startCol = -1;

            ws.eachRow((row, rowNum) => {
                if (startRow !== -1) return;
                row.eachCell((cell, colNum) => {
                    const v = cell.value ? cell.value.toString().toUpperCase() : '';
                    if (v.includes(startLabelSearch)) {
                        startRow = rowNum;
                        startCol = colNum;
                    }
                });
            });

            if (startRow === -1) return;

            // School Info
            ws.getRow(startRow).getCell(startCol + 4).value = semData.school;
            ws.getRow(startRow).getCell(startCol + 27).value = semData.schoolId;
            ws.getRow(startRow).getCell(startCol + 38).value = semData.gradeLevel;
            ws.getRow(startRow).getCell(startCol + 50).value = semData.sy;
            ws.getRow(startRow).getCell(startCol + 59).value = semData.sem;

            // Find "SUBJECTS" header
            let tableRow = -1;
            for (let r = startRow + 1; r < startRow + 10; r++) {
                const rowText = ws.getRow(r).values.join(' ').toUpperCase();
                if (rowText.includes('SUBJECTS')) {
                    tableRow = r + 1;
                    break;
                }
            }

            if (tableRow !== -1) {
                semData.subjects.forEach((subj, i) => {
                    const r = tableRow + i;
                    const row = ws.getRow(r);
                    row.getCell(1).value = subj.type;
                    row.getCell(9).value = subj.subject;
                    row.getCell(45).value = subj.q1;
                    row.getCell(50).value = subj.q2;
                    row.getCell(55).value = subj.final;
                    row.getCell(60).value = subj.action;
                });

                // General Average
                for (let r = tableRow + 10; r < tableRow + 20; r++) {
                    const rowStr = ws.getRow(r).values.join(' ').toUpperCase();
                    if (rowStr.includes('GENERAL AVE')) {
                        ws.getRow(r).getCell(55).value = semData.genAve;
                        break;
                    }
                }
            }
        };

        fillSemester(front, data.semester1, 'SCHOOL:');
        let schoolOccurrences = 0;
        front.eachRow((row, rowNum) => {
            row.eachCell((cell, colNum) => {
                if (cell.value && cell.value.toString().toUpperCase() === 'SCHOOL:') {
                    schoolOccurrences++;
                    if (schoolOccurrences === 2) {
                        const semData = data.semester2;
                        row.getCell(colNum + 4).value = semData.school;
                        row.getCell(colNum + 27).value = semData.schoolId;
                    }
                }
            });
        });

        fillSemester(back, data.semester3, 'SCHOOL:');
        let backOccurrences = 0;
        back.eachRow((row, rowNum) => {
            row.eachCell((cell, colNum) => {
                if (cell.value && cell.value.toString().toUpperCase() === 'SCHOOL:') {
                    backOccurrences++;
                    if (backOccurrences === 2) {
                        const semData = data.semester4;
                        row.getCell(colNum + 4).value = semData.school;
                        row.getCell(colNum + 27).value = semData.schoolId;
                    }
                }
            });
        });

        // Certification Fallback
        writeToLabel(back, 'Track/Strand', data.certification.trackStrand, 3);
        writeToLabel(back, 'SHS General Average', data.certification.genAve, 3);
        writeToLabel(back, 'Date of Graduation', data.certification.gradDate, 2, 2);
        writeToLabel(back, 'Name of School', data.certification.schoolName, 2);

        // Generate Blob and download
        const uint8Array = await wb.xlsx.writeBuffer();
        const blob = new Blob([uint8Array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        const studentName = (data.info.lname || 'Form137').replace(/[^a-z0-9]/gi, '_');
        a.download = `Form137_${studentName}_${Date.now()}.xlsx`;

        document.body.appendChild(a);
        a.click();

        setTimeout(() => {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 100);

        return { success: true };
    } catch (error) {
        console.error('Error generating Excel file:', error);
        return { success: false, error: error.message };
    }
}
