document.addEventListener('DOMContentLoaded', () => {
    const printContent = document.getElementById('print-content');
    const btnPrint = document.getElementById('do-print');
    const btnClose = document.getElementById('do-close');

    // Load data from localStorage
    const rawData = localStorage.getItem('printStudentData');
    if (!rawData) {
        printContent.innerHTML = '<div style="padding: 20px; text-align: center; color: red;">No student data found for printing.</div>';
        return;
    }

    const d = JSON.parse(rawData);
    renderPrintForm(d, printContent);

    // Event Listeners
    if (btnPrint) {
        btnPrint.addEventListener('click', () => {
            window.print();
        });
    }

    if (btnClose) {
        btnClose.addEventListener('click', () => {
            window.close();
        });
    }
});

function renderPrintForm(d, container) {
    // Ensure all semesters exist
    if (!d.semester3) d.semester3 = d.semester1; // Fallback structure if missing
    if (!d.semester4) d.semester4 = d.semester1;

    // === PAGE 1 ===
    const page1 = document.createElement('div');
    page1.className = 'paper';
    page1.innerHTML = `
        <table style="width:100%; border-collapse:collapse; margin-bottom:2px;">
            <tr>
                <td width="40" style="border:none; padding:0; vertical-align:middle;"><img src="assets/images/school_logo.webp" width="38"></td>
                <td style="border:none; text-align:center; font-size:5.5pt; line-height:1.1; padding:0 5px; vertical-align:middle;">
                    <div>REPUBLIC OF THE PHILIPPINES</div>
                    <div style="font-weight:bold;">DEPARTMENT OF EDUCATION</div>
                    <div style="font-weight:bold; font-size:6.5pt;">SENIOR HIGH SCHOOL STUDENT PERMANENT RECORD</div>
                </td>
                <td width="55" style="border:none; text-align:right; padding:0; vertical-align:middle;">
                    <div style="font-size:4.5pt; font-weight:bold; margin-bottom:1px;">FORM137-SHS</div>
                    <div style="text-align:right;"><img src="assets/images/deped_logo.webp" width="38"></div>
                </td>
            </tr>
        </table>

        <div class="hdr">LEARNER'S INFORMATION</div>
        <table class="clean-tbl" style="width:100%;">
            <tr>
                <td style="width:33%; font-size:6pt; font-weight:bold;">LAST NAME: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 70px); font-weight:normal;"><input class="inp" value="${d.info.lname}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:34%; font-size:6pt; font-weight:bold;">FIRST NAME: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 75px); font-weight:normal;"><input class="inp" value="${d.info.fname}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:33%; font-size:6pt; font-weight:bold;">MIDDLE NAME: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 90px); font-weight:normal;"><input class="inp" value="${d.info.mname}" disabled style="width:100%; border:none;"></span></td>
            </tr>
            <tr>
                <td style="font-size:6pt; font-weight:bold;">LRN: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 35px); font-weight:normal;"><input class="inp" value="${d.info.lrn}" disabled style="width:100%; border:none;"></span></td>
                <td colspan="2" style="font-size:6pt; font-weight:bold;">Date of Birth (mm/dd/yyyy): <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 145px); font-weight:normal;"><input class="inp" value="${d.info.birthdate}" disabled style="width:100%; border:none;"></span></td>
            </tr>
            <tr>
                <td style="font-size:6pt; font-weight:bold;">Sex: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 30px); font-weight:normal;"><input class="inp" value="${d.info.sex}" disabled style="width:100%; border:none;"></span></td>
                <td colspan="2" style="font-size:6pt; font-weight:bold;">Date of SHS Admission (mm/dd/yyyy): <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 195px); font-weight:normal;"><input class="inp" value="${d.info.admissionDate}" disabled style="width:100%; border:none;"></span></td>
            </tr>
        </table>

        <div class="hdr">ELIGIBILITY FOR SHS ENROLMENT</div>
        <table style="width:100%; border:none; border-collapse:collapse; font-size:5.5pt;">
            <tr>
                <td style="width:18%; border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.hsCompleter ? 'checked' : ''} disabled> High School Completer*</td>
                <td style="width:14%; border:none; padding:1px 2px; white-space:nowrap;">Gen. Ave: <input class="inp" value="${d.eligibility.hsGenAve}" disabled style="border:none; border-bottom:1px solid #000; width:50%; display:inline-block;"></td>
                <td style="width:18%; border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.jhsCompleter ? 'checked' : ''} disabled> JHS Completer</td>
                <td style="width:14%; border:none; padding:1px 2px; white-space:nowrap;">Gen. Ave: <input class="inp" value="${d.eligibility.jhsGenAve}" disabled style="border:none; border-bottom:1px solid #000; width:50%; display:inline-block;"></td>
                <td style="width:36%; border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.others ? 'checked' : ''} disabled> Others: <input class="inp" value="${d.eligibility.othersSpec}" disabled style="border:none; border-bottom:1px solid #000; width:calc(100% - 65px); display:inline-block;"></td>
            </tr>
            <tr>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">Date of Graduation: <input class="inp" value="${d.eligibility.gradDate}" disabled style="border:none; border-bottom:1px solid #000; width:55%; display:inline-block;"></td>
                <td style="border:none; padding:1px 2px; white-space:nowrap;">Name of School: <input class="inp" value="${d.eligibility.schoolName}" disabled style="border:none; border-bottom:1px solid #000; width:calc(100% - 90px); display:inline-block;"></td>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">School Address: <input class="inp" value="${d.eligibility.schoolAddress}" disabled style="border:none; border-bottom:1px solid #000; width:calc(100% - 100px); display:inline-block;"></td>
            </tr>
            <tr>
                <td style="border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.pept ? 'checked' : ''} disabled> PEPT Passer**</td>
                <td style="border:none; padding:1px 2px; white-space:nowrap;">Rating: <input class="inp" value="${d.eligibility.peptRating}" disabled style="border:none; border-bottom:1px solid #000; width:60%; display:inline-block;"></td>
                <td style="border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.als ? 'checked' : ''} disabled> ALS A&E***</td>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">Rating: <input class="inp" value="${d.eligibility.alsRating}" disabled style="border:none; border-bottom:1px solid #000; width:calc(100% - 55px); display:inline-block;"></td>
            </tr>
            <tr>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">Date of Exam: <input class="inp" value="${d.eligibility.examDate}" disabled style="border:none; border-bottom:1px solid #000; width:55%; display:inline-block;"></td>
                <td colspan="3" style="border:none; padding:1px 2px; white-space:nowrap;">CLC Name and Address: <input class="inp" value="${d.eligibility.clcName}" disabled style="border:none; border-bottom:1px solid #000; width:calc(100% - 135px); display:inline-block;"></td>
            </tr>
            <tr style="font-size:5pt; font-style:italic;">
                <td colspan="2" style="border:none; padding:1px 2px;">*HS Completers graduated under old curriculum</td>
                <td colspan="3" style="border:none; padding:1px 2px;">***ALS A&E - Alternative Learning System Accred. & Equiv. Test</td>
            </tr>
            <tr style="font-size:5pt; font-style:italic;">
                <td colspan="5" style="border:none; padding:1px 2px;">**PEPT - Philippine Educational Placement Test</td>
            </tr>
        </table>

        ${renderSemester(d.semester1)}
        ${renderSemester(d.semester2)}
    `;
    container.appendChild(page1);

    // === PAGE 2 ===
    const page2 = document.createElement('div');
    page2.className = 'paper';
    page2.innerHTML = `
        <div style="text-align:right; font-size:6pt; margin-bottom:3px;">Page 2 &nbsp; Form 137-SHS</div>
        ${renderSemester(d.semester3)}
        ${renderSemester(d.semester4)}
        
        <div style="font-size:6pt; margin-top:3px;">
            <div style="margin-bottom:2px;">
                <span style="font-weight:bold;">Track/Strand Accomplished:</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:340px; margin:0 20px 0 5px; vertical-align:bottom;">
                    <input class="inp" value="${d.certification.trackStrand}" disabled style="width:100%; border:none; background:transparent; padding:0;">
                </span>
                <span style="font-weight:bold;">SHS General Average:</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:120px; margin-left:5px; vertical-align:bottom;">
                    <input class="inp" value="${d.certification.genAve}" disabled style="width:100%; border:none; background:transparent; padding:0;">
                </span>
            </div>
            <div>
                <span style="font-weight:bold;">Awards/Honors Received:</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:318px; margin:0 20px 0 5px; vertical-align:bottom;">
                    <input class="inp" value="${d.certification.awards}" disabled style="width:100%; border:none; background:transparent; padding:0;">
                </span>
                <span style="font-weight:bold;">Date of SHS Graduation (MM/DD/YYYY):</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:90px; margin-left:5px; vertical-align:bottom;">
                    <input type="date" class="inp" value="${d.certification.gradDate}" disabled style="width:100%; border:none; background:transparent; padding:0;">
                </span>
            </div>
        </div>

        <table style="width:100%; border:none; border-collapse:collapse; margin-top:3px;">
            <tr>
                <td style="border:none; padding:3px 5px; font-size:6pt; font-weight:bold; width:73%; border-bottom:1px solid #000;">
                    Certified by:
                </td>
                <td style="border:none; border-left:1px solid #000; border-bottom:1px solid #000; padding:3px 5px; font-size:5.5pt; width:27%; text-align:right;">
                    <span style="margin-right:5px;">Date:</span>
                    <span style="border-bottom:1px solid #000; display:inline-block; width:100px; vertical-align:bottom;">
                        <input type="date" class="inp" value="${d.certification.certDate}" disabled style="width:100%; border:none; background:transparent; text-align:center; padding:0; font-size:5.5pt;">
                    </span>
                </td>
            </tr>
            <tr>
                <td style="border:none; padding:5px 5px 2px 5px; font-size:6pt; border-bottom:1px solid #000;">
                    <div style="width:350px; border-bottom:1px solid #000; margin:0 0 3px 0; height:30px;"></div>
                    <div style="font-size:5.5pt;">
                        Signature of School Head over Printed Name
                    </div>
                </td>
                <td style="border:none; border-left:1px solid #000; border-bottom:1px solid #000;"></td>
            </tr>
            <tr>
                <td style="border:none; padding:3px 5px; font-size:6pt; vertical-align:top; border-bottom:1px solid #000;">
                    <div style="border:1px solid #000; padding:5px; margin:2px 0;">
                        <div style="font-weight:bold; margin-bottom:3px;">NOTE:</div>
                        <div style="font-size:5.5pt; font-style:italic; line-height:1.3;">
                            This permanent record or a photocopy of this permanent record that bears the seal of the school and the original signature in ink of the School Head shall be considered valid for all legal purposes. Any erasure or alteration made on this copy should be validated by the School Head.<br>
                            If the student transfers to another school, the originating school should produce one (1) certified true copy of this permanent record for safekeeping. The receiving school shall continue filling up the original form.<br>
                            Upon graduation, the school from which the student graduated should keep the original form and produce one (1) certified true copy for the Division Office.
                        </div>
                    </div>
                </td>
                <td style="border:none; border-left:1px solid #000; border-bottom:1px solid #000; padding:3px 5px; font-size:5.5pt; vertical-align:top; text-align:right;">
                    Place School Seal Here:
                </td>
            </tr>
            <tr>
                <td style="border:none; padding:3px 5px; font-size:6pt; vertical-align:top; border-bottom:1px solid #000;">
                    <div style="font-weight:bold; margin-bottom:3px;">REMARKS: <span style="font-weight:normal; font-style:italic;">(Please indicate the purpose for which this permanent record will be used)</span></div>
                    <textarea class="inp" style="width:100%; border:none; background:transparent; font-family:inherit; resize:none; min-height:40px; font-size:6pt; padding:0; line-height:1.3; margin-bottom:5px;" disabled>${d.certification.remarks}</textarea>
                    <div style="margin-top:5px;">
                        <span style="font-weight:bold; margin-right:5px;">Date Issued (MM/DD/YYYY):</span>
                        <span style="border-bottom:1px solid #000; display:inline-block; width:120px; vertical-align:bottom;">
                            <input type="date" class="inp" value="${d.certification.dateIssued}" disabled style="width:100%; border:none; background:transparent; padding:0; font-size:6pt;">
                        </span>
                    </div>
                </td>
                <td style="border:none; border-left:1px solid #000; border-bottom:1px solid #000;"></td>
            </tr>
        </table>
    `;
    container.appendChild(page2);
}

function renderSemester(sem) {
    if (!sem) return '';
    return `
        <div class="hdr" style="margin-top:5px;">SCHOLASTIC RECORD</div>
        <table class="clean-tbl" style="width:100%; font-size:6pt;">
            <tr>
                <td style="width:35%; font-weight:bold;">SCHOOL: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 55px); font-weight:normal;"><input class="inp" value="${sem.school}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:15%; font-weight:bold;">SCHOOL ID: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 75px); font-weight:normal;"><input class="inp" value="${sem.schoolId}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:15%; font-weight:bold;">GRADE LEVEL: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 90px); font-weight:normal;"><input class="inp" value="${sem.gradeLevel}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:15%; font-weight:bold;">SY: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 25px); font-weight:normal;"><input class="inp" value="${sem.sy}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:20%; font-weight:bold;">SEM: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 35px); font-weight:normal;"><input class="inp" value="${sem.sem}" disabled style="width:100%; border:none;"></span></td>
            </tr>
            <tr>
                <td colspan="3" style="font-weight:bold;">TRACK/STRAND: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 110px); font-weight:normal;"><input class="inp" value="${sem.trackStrand}" disabled style="width:100%; border:none;"></span></td>
                <td colspan="2" style="font-weight:bold;">SECTION: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 60px); font-weight:normal;"><input class="inp" value="${sem.section}" disabled style="width:100%; border:none;"></span></td>
            </tr>
        </table>

        <table class="form" style="font-size:5.5pt;">
            <thead>
                <tr>
                    <th rowspan="2" width="15%" style="font-size:5pt;">CORE/APPLIED/<br>SPECIALIZED</th>
                    <th rowspan="2" width="42%">SUBJECTS</th>
                    <th colspan="2" width="14%">Quarter</th>
                    <th rowspan="2" width="11%">SEM FINAL</th>
                    <th rowspan="2" width="12%">ACTION TAKEN</th>
                </tr>
                <tr>
                    <th width="7%">${sem.q1Label || '1ST'}</th>
                    <th width="7%">${sem.q2Label || '2ND'}</th>
                </tr>
            </thead>
            <tbody>
                ${sem.subjects.slice(0, 9).map((s) => `
                    <tr>
                        <td><input class="inp c" value="${s.type}" disabled style="font-size:5pt;"></td>
                        <td><input class="inp" value="${s.subject}" disabled></td>
                        <td><input class="inp c" value="${s.q1}" disabled></td>
                        <td><input class="inp c" value="${s.q2}" disabled></td>
                        <td><input class="inp c" value="${s.final}" disabled></td>
                        <td><input class="inp c" value="${s.action}" disabled></td>
                    </tr>
                `).join('')}
                <tr>
                    <td colspan="4" style="text-align:right; font-weight:bold; padding-right:5px;">General Ave. for the Semester:</td>
                    <td><input class="inp c" value="${sem.genAve}" disabled style="border:none; width:100%;"></td>
                    <td></td>
                </tr>
            </tbody>
        </table>

        <div style="font-size:6pt; font-weight:bold; margin-top:2px; display:flex; align-items:flex-end;">
            <span style="margin-right:5px;">REMARKS:</span>
            <div style="flex-grow:1; border-bottom:1px solid #000;">
                <input class="inp" value="${sem.remarks}" disabled style="width:100%; border:none; background:transparent;">
            </div>
        </div>

        <table class="clean-tbl" style="font-size:5.5pt; width:100%;">
            <tr>
                <td style="width:33%;">Prepared by: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 80px);"><input class="inp" value="${sem.adviserName}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:34%;">Certified: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 65px);"><input class="inp" value="${sem.certName}" disabled style="width:100%; border:none;"></span></td>
                <td style="width:33%;">Date Checked: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 90px);"><input class="inp" value="${sem.dateChecked}" disabled style="width:100%; border:none;"></span></td>
            </tr>
        </table>
        <table class="clean-tbl" style="font-size:5pt; font-style:italic; width:100%; margin-top:1px; border-collapse:separate; border-spacing:10px 0;">
            <tr>
                <td style="width:50%; text-align:center; border-bottom:1px solid #000; padding-bottom:1px;">&nbsp;</td>
                <td style="width:50%; text-align:center; border-bottom:1px solid #000; padding-bottom:1px;">&nbsp;</td>
            </tr>
            <tr>
                <td style="text-align:center; padding-top:1px;">Signature of Adviser over Printed Name</td>
                <td style="text-align:center; padding-top:1px;">Signature of Authorized Person over Printed Name</td>
            </tr>
        </table>

        <div class="hdr" style="background:#f0f0f0; margin-top:5px;">REMEDIAL CLASSES</div>
        <table class="clean-tbl" style="font-size:5.5pt; width:100%;">
            <tr>
                <td colspan="2">Conducted from: <input class="inp" value="${sem.remedial.from}" disabled style="border-bottom:1px solid #000; width:50px; display:inline-block; border:none;"></td>
                <td>to: <input class="inp" value="${sem.remedial.to}" disabled style="border-bottom:1px solid #000; width:50px; display:inline-block; border:none;"></td>
                <td style="font-weight:bold;">SCHOOL: <input class="inp" value="${sem.remedial.school}" disabled style="border-bottom:1px solid #000; width:calc(100% - 60px); display:inline-block; border:none;"></td>
                <td colspan="2" style="font-weight:bold;">SCHOOL ID: <input class="inp" value="${sem.remedial.schoolId}" disabled style="border-bottom:1px solid #000; width:calc(100% - 75px); display:inline-block; border:none;"></td>
            </tr>
        </table>
        <table class="form" style="font-size:5.5pt;">
            <thead>
                <tr>
                    <th width="15%">CORE/APPLIED/SPECIALIZED</th>
                    <th width="39%">SUBJECTS</th>
                    <th width="11%">SEM FINAL</th>
                    <th width="12%">REMEDIAL MARK</th>
                    <th width="13%">RECOMPUTED</th>
                    <th width="10%">ACTION</th>
                </tr>
            </thead>
            <tbody>
                ${sem.remedial.subjects.slice(0, 4).map((rs) => `
                    <tr>
                        <td><input class="inp c" value="${rs.type}" disabled></td>
                        <td><input class="inp" value="${rs.subject}" disabled></td>
                        <td><input class="inp c" value="${rs.semGrade}" disabled></td>
                        <td><input class="inp c" value="${rs.remedialMark}" disabled></td>
                        <td><input class="inp c" value="${rs.recomputedGrade}" disabled></td>
                        <td><input class="inp c" value="${rs.action}" disabled></td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
        <table class="form" style="font-size:5.5pt;">
            <tr><td width="50%">Name of Teacher/Adviser: <input class="inp" value="${sem.remedial.teacherName}" disabled style="border-bottom:1px solid #000; width:calc(100% - 145px);"></td><td>Signature:</td></tr>
        </table>
    `;
}
