// excel-generator-worker.js
// Uses exceljs to generate/parse files by searching for labels like "LAST NAME:"

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

process.on('message', async (msg) => {
    if (msg.type === 'generate-print-file') {
        const { data, templatePath } = msg;

        try {
            if (!fs.existsSync(templatePath)) throw new Error('Template not found');

            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(templatePath);

            const front = wb.getWorksheet('FRONT');
            const back = wb.getWorksheet('BACK');

            if (!front || !back) throw new Error('FRONT or BACK worksheet missing in template');

            // --- HELPER: Find label and write relative to it ---
            const writeToLabel = (ws, label, value, offsetCol = 1, offsetRow = 0) => {
                if (value === undefined || value === null) return;
                const search = label.toUpperCase();
                ws.eachRow((row, rowNum) => {
                    row.eachCell((cell, colNum) => {
                        const cellVal = cell.value ? cell.value.toString().toUpperCase() : '';
                        if (cellVal.includes(search)) {
                            const targetRow = ws.getRow(rowNum + offsetRow);
                            const targetCell = targetRow.getCell(colNum + offsetCol);
                            targetCell.value = value;
                        }
                    });
                });
            };

            // --- FILL FRONT SHEET ---
            // Learner Information
            writeToLabel(front, 'LAST NAME:', data.info.lname);
            writeToLabel(front, 'FIRST NAME:', data.info.fname);
            writeToLabel(front, 'MIDDLE NAME:', data.info.mname);
            writeToLabel(front, 'LRN:', data.info.lrn);
            writeToLabel(front, 'SEX:', data.info.sex);
            writeToLabel(front, 'DATE OF BIRTH', data.info.birthdate, 3);
            writeToLabel(front, 'DATE OF SHS ADMISSION', data.info.admissionDate, 5);

            // Eligibility
            if (data.eligibility.hsCompleter) writeToLabel(front, 'High School Completer', 'X', -1);
            if (data.eligibility.jhsCompleter) writeToLabel(front, 'Junior High School Completer', 'X', -1);
            writeToLabel(front, 'Gen. Ave:', data.eligibility.hsGenAve);
            writeToLabel(front, 'Date of Graduation', data.eligibility.gradDate, 3);
            writeToLabel(front, 'Name of School:', data.eligibility.schoolName, 2);
            writeToLabel(front, 'School Address:', data.eligibility.schoolAddress, 1);

            // --- FILL SEMESTERS ---
            // Heuristic mapping based on common SF10 layouts
            const fillSemester = (ws, semData, startLabelSearch, isRemedial = false) => {
                if (!semData) return;
                let startRow = -1;
                let startCol = -1;

                // Find the specific semester block
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

                // Find "SUBJECTS" header to start table
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
                        // Using common column indexes for SF10 templates
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

            // Filling by searching for unique markers
            // Note: Template might use "1st Semester" or "FIRST SEMESTER"
            fillSemester(front, data.semester1, 'SCHOOL:'); // First occurrence on FRONT
            // For 2nd sem on FRONT, we need to find the SECOND "SCHOOL:"
            let schoolOccurrences = 0;
            front.eachRow((row, rowNum) => {
                row.eachCell((cell, colNum) => {
                    if (cell.value && cell.value.toString().toUpperCase() === 'SCHOOL:') {
                        schoolOccurrences++;
                        if (schoolOccurrences === 2) {
                            // Manual fill for 2nd occurrence
                            const startRow = rowNum;
                            const semData = data.semester2;
                            row.getCell(colNum + 4).value = semData.school;
                            row.getCell(colNum + 27).value = semData.schoolId;
                            // ... similarly for others if needed ...
                        }
                    }
                });
            });

            // --- FILL SEMESTERS (ALL 4) ---
            const fillSem = (ws, semData, instance = 1) => {
                if (!semData || !semData.subjects.length) return;
                let currentInstance = 0;
                let startRow = -1;
                let startCol = -1;

                ws.eachRow((row, rowNum) => {
                    if (startRow !== -1) return;
                    row.eachCell((cell, colNum) => {
                        const v = cell.value ? cell.value.toString().toUpperCase() : '';
                        if (v === 'SCHOOL:') {
                            currentInstance++;
                            if (currentInstance === instance) {
                                startRow = rowNum;
                                startCol = colNum;
                            }
                        }
                    });
                });

                if (startRow === -1) return;

                // School Info (Fixed offsets from the "SCHOOL:" label)
                ws.getRow(startRow).getCell(startCol + 4).value = semData.school;
                ws.getRow(startRow).getCell(startCol + 27).value = semData.schoolId;
                ws.getRow(startRow).getCell(startCol + 38).value = semData.gradeLevel;
                ws.getRow(startRow).getCell(startCol + 50).value = semData.sy;
                ws.getRow(startRow).getCell(startCol + 59).value = semData.sem;

                // Additional Info (Track/Section)
                // These are usually 1 or 2 rows below or on the same line
                ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
                    if (rowNum >= startRow && rowNum <= startRow + 3) {
                        row.eachCell((cell, colNum) => {
                            const v = cell.value ? cell.value.toString().toUpperCase() : '';
                            if (v.includes('TRACK')) ws.getRow(rowNum).getCell(colNum + 6).value = semData.trackStrand;
                            if (v.includes('SECTION')) ws.getRow(rowNum).getCell(colNum + 6).value = semData.section;
                        });
                    }
                });

                // Find "SUBJECTS" header below this block
                let tableRow = -1;
                for (let r = startRow + 1; r < startRow + 10; r++) {
                    const rowText = (ws.getRow(r).values || []).join(' ').toUpperCase();
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
                    for (let r = tableRow + 10; r < tableRow + 25; r++) {
                        const rowStr = (ws.getRow(r).values || []).join(' ').toUpperCase();
                        if (rowStr.includes('GENERAL AVE')) {
                            ws.getRow(r).getCell(55).value = semData.genAve;
                            break;
                        }
                    }
                }
            };

            // Mapping:
            // FRONT: Inst 1 = Sem 1, Inst 2 = Sem 2
            // BACK: Inst 1 = Sem 3, Inst 2 = Sem 4
            fillSem(front, data.semester1, 1);
            fillSem(front, data.semester2, 2);
            fillSem(back, data.semester3, 1);
            fillSem(back, data.semester4, 2);

            // Certification (BACK sheet)
            writeToLabel(back, 'CERTIFICATION', 'X', -1); // Just a marker
            // Map the certification fields if data exists
            if (data.certification) {
                writeToLabel(back, 'Date of Graduation', data.certification.gradDate, 3);
                writeToLabel(back, 'Name of School', data.certification.schoolName, 2);
            }

            const tempDir = require('os').tmpdir();
            const studentName = (data.info.lname || 'Form137').replace(/[^a-z0-9]/gi, '_');
            const outputPath = path.join(tempDir, `${studentName}_${Date.now()}.xlsx`);

            await wb.xlsx.writeFile(outputPath);
            process.send({ type: 'result', success: true, filePath: outputPath });

        } catch (e) {
            console.error(e);
            process.send({ type: 'result', success: false, error: e.message });
        }
    }

    if (msg.type === 'parse-excel') {
        // Implementation for parsing back to JSON if needed
        process.send({ type: 'parse-result', success: false, error: 'Not implemented' });
    }
});
