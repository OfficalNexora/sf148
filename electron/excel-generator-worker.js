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
            const annexSheet = wb.getWorksheet('ANNEX');

            if (!front || !back) throw new Error('FRONT or BACK worksheet missing in template');

            // --- HELPER: Find label and write relative to it (Fallback) ---
            const writeToLabel = (ws, label, value, offsetCol = 1, offsetRow = 0) => {
                if (value === undefined || value === null || value === '') return;
                const search = label.toUpperCase();
                ws.eachRow((row, rowNum) => {
                    row.eachCell((cell, colNum) => {
                        const cellVal = cell.value ? cell.value.toString().toUpperCase() : '';
                        if (cellVal.includes(search)) {
                            // Only write if the target cell is EMPTY (to avoid overwriting if placeholders already did it)
                            const targetRow = ws.getRow(rowNum + offsetRow);
                            const targetCell = targetRow.getCell(colNum + offsetCol);
                            if (!targetCell.value) targetCell.value = value;
                        }
                    });
                });
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
                // Compatibility Aliases (from PLACEHOLDER(ALL).xlsx)
                'lastname': data.info.lname,
                'firstname': data.info.fname,
                'middlename': data.info.mname,
                'LRN': data.info.lrn,
                'date of birth': data.info.birthdate,
                'dateofadmission': data.info.admissionDate,
                'genave': data.semester1.genAve, // General placeholder for first sem average
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
                    // Remedial for each semester
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

            // Dynamic Annex Placeholders (Max 36 rows)
            if (data.annex) {
                data.annex.forEach((subj, i) => {
                    const idx = i + 1;
                    placeholderMap[`asub_${idx}`] = subj.subject;
                    placeholderMap[`atype_${idx}`] = subj.type;
                    placeholderMap[`aq1_${idx}`] = subj.q1;
                    placeholderMap[`aq2_${idx}`] = subj.q2;
                    placeholderMap[`afin_${idx}`] = subj.final;
                    placeholderMap[`aact_${idx}`] = subj.action;
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

            replacePlaceholders(front);
            replacePlaceholders(back);
            if (annexSheet) replacePlaceholders(annexSheet);

            // --- FILL FRONT SHEET (Fallback) ---
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

            const fillAnnex = (ws, annexData) => {
                if (!ws || !annexData || !annexData.length) return;

                const findSectionRange = (headerText) => {
                    let start = -1;
                    ws.eachRow((row, rowNum) => {
                        if (start !== -1) return;
                        const cells = Array.isArray(row.values) ? row.values.join(' ').toUpperCase() : '';
                        if (cells.includes(headerText.toUpperCase())) {
                            start = rowNum;
                        }
                    });
                    return start;
                };

                const coreStart = findSectionRange('CORE SUBJECTS');
                const appliedStart = findSectionRange('APPLIED SUBJECTS');
                const specializedStart = findSectionRange('SPECIALIZED SUBJECTS');
                const otherStart = findSectionRange('OTHER SUBJECTS');

                const markCheckbox = (row, isPassed) => {
                    if (isPassed) {
                        // Based on template, Col 1 or 2 is the checkbox
                        const cell1 = row.getCell(1);
                        const cell2 = row.getCell(2);
                        // Write to whichever looks like a checkbox or just Col 1
                        cell1.value = '/';
                    }
                };

                // Core Subjects (Indices 0-14)
                if (coreStart !== -1) {
                    const dataRange = annexData.slice(0, 15);
                    dataRange.forEach(subj => {
                        if (!subj.subject) return;
                        // Search for the subject in the rows below the header
                        const endRow = appliedStart !== -1 ? appliedStart : coreStart + 25;
                        for (let r = coreStart + 1; r < endRow; r++) {
                            const row = ws.getRow(r);
                            const rowText = (row.values || []).join(' ').toUpperCase();
                            if (rowText.includes(subj.subject.toUpperCase())) {
                                markCheckbox(row, subj.action === 'PASSED');
                                break;
                            }
                        }
                    });
                }

                // Applied Subjects (Indices 15-21)
                if (appliedStart !== -1) {
                    const dataRange = annexData.slice(15, 22);
                    dataRange.forEach(subj => {
                        if (!subj.subject) return;
                        const endRow = specializedStart !== -1 ? specializedStart : appliedStart + 15;
                        for (let r = appliedStart + 1; r < endRow; r++) {
                            const row = ws.getRow(r);
                            const rowText = (row.values || []).join(' ').toUpperCase();
                            if (rowText.includes(subj.subject.toUpperCase())) {
                                markCheckbox(row, subj.action === 'PASSED');
                                break;
                            }
                        }
                    });
                }

                // Specialized Subjects (Indices 22-30) - Blank lines
                if (specializedStart !== -1) {
                    const dataRange = annexData.slice(22, 31);
                    let currentRow = specializedStart + 1;
                    dataRange.forEach(subj => {
                        if (!subj.subject) return;
                        const endRow = otherStart !== -1 ? otherStart : specializedStart + 15;
                        while (currentRow < endRow) {
                            const row = ws.getRow(currentRow);
                            const cell9 = row.getCell(9); // Usual subject column
                            if (!cell9.value || cell9.value.toString().includes('PlaceHolder')) {
                                cell9.value = subj.subject;
                                markCheckbox(row, subj.action === 'PASSED');
                                currentRow++;
                                break;
                            }
                            currentRow++;
                        }
                    });
                }

                // Other Subjects (Indices 31-35) - Blank lines
                if (otherStart !== -1) {
                    const dataRange = annexData.slice(31, 36);
                    let currentRow = otherStart + 1;
                    dataRange.forEach(subj => {
                        if (!subj.subject) return;
                        while (currentRow < otherStart + 10) {
                            const row = ws.getRow(currentRow);
                            const cell9 = row.getCell(9);
                            if (!cell9.value) {
                                cell9.value = subj.subject;
                                markCheckbox(row, subj.action === 'PASSED');
                                currentRow++;
                                break;
                            }
                            currentRow++;
                        }
                    });
                }
            };

            // Mapping:
            // FRONT: Inst 1 = Sem 1, Inst 2 = Sem 2
            // BACK: Inst 1 = Sem 3, Inst 2 = Sem 4
            fillSem(front, data.semester1, 1);
            fillSem(front, data.semester2, 2);
            fillSem(back, data.semester3, 1);
            fillSem(back, data.semester4, 2);

            if (annexSheet) {
                fillAnnex(annexSheet, data.annex);
            }

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
