import React from 'react';

/**
 * Form137Preview — Renders a faithful reproduction of the DepEd SF10-SHS
 * (Form 137 - Senior High School Student Permanent Record) directly in HTML.
 *
 * Layout derived from template-analysis.txt field mapping:
 * FRONT: Header, Learner Info, Eligibility, Semester Block 1, Semester Block 2
 * BACK:  Semester Block 3, Semester Block 4, Certification Note
 */
export default function Form137Preview({ data }) {
    const info = data.info || {};
    const elig = data.eligibility || {};
    const sem1 = data.semester1 || {};
    const sem2 = data.semester2 || {};
    const cert = data.certification || {};

    // ── Shared styles ──
    const page = {
        width: '100%',
        background: '#fff',
        padding: '24px 28px',
        fontFamily: "'Arial Narrow', Arial, sans-serif",
        fontSize: '8pt',
        color: '#000',
        lineHeight: 1.25,
        marginBottom: '24px',
        boxSizing: 'border-box',
    };
    const tbl = { width: '100%', borderCollapse: 'collapse', border: '1px solid #000', tableLayout: 'fixed' };
    const c = { border: '1px solid #000', padding: '1px 3px', fontSize: '7.5pt', verticalAlign: 'middle' };
    const cl = { ...c, fontWeight: 'bold', whiteSpace: 'nowrap' };
    const cv = { ...c, minWidth: '40px' };
    const hdr = {
        background: '#d9d9d9',
        fontWeight: 'bold',
        textAlign: 'center',
        fontSize: '8pt',
        border: '1px solid #000',
        padding: '2px 4px',
        letterSpacing: '0.5px',
    };
    const subHdr = { ...hdr, fontSize: '7pt', background: '#e6e6e6' };

    // ── Semester Block ──
    const renderSemBlock = (sem, label) => {
        const subjects = sem.subjects || [];
        const remedial = sem.remedial || {};
        const remSubjects = remedial.subjects || [];
        const SUBJECT_ROWS = 9;

        return (
            <div style={{ marginBottom: '6px' }}>
                {/* School / Grade / SY Info */}
                <table style={tbl}>
                    <colgroup>
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '25%' }} />
                        <col style={{ width: '10%' }} />
                        <col style={{ width: '10%' }} />
                        <col style={{ width: '12%' }} />
                        <col style={{ width: '9%' }} />
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '10%' }} />
                        <col style={{ width: '4%' }} />
                        <col style={{ width: '4%' }} />
                    </colgroup>
                    <tbody>
                        <tr>
                            <td style={cl}>SCHOOL:</td>
                            <td style={cv} colSpan="2">{sem.school}</td>
                            <td style={cl}>SCHOOL ID:</td>
                            <td style={cv}>{sem.schoolId}</td>
                            <td style={cl}>GRADE LEVEL:</td>
                            <td style={cv}>{sem.gradeLevel}</td>
                            <td style={cl}>SY:</td>
                            <td style={cv}>{sem.sy}</td>
                            <td style={{ ...cl, fontSize: '6.5pt' }}>SEM: {sem.sem}</td>
                        </tr>
                        <tr>
                            <td style={cl} colSpan="2">TRACK/STRAND: <span style={{ fontWeight: 'normal' }}>{sem.trackStrand}</span></td>
                            <td style={cl} colSpan="8">SECTION: <span style={{ fontWeight: 'normal' }}>{sem.section}</span></td>
                        </tr>
                    </tbody>
                </table>

                {/* Scholastic Record Header */}
                <div style={hdr}>SCHOLASTIC RECORD</div>

                {/* Subjects Table */}
                <table style={tbl}>
                    <colgroup>
                        <col style={{ width: '22%' }} />
                        <col style={{ width: '33%' }} />
                        <col style={{ width: '10%' }} />
                        <col style={{ width: '10%' }} />
                        <col style={{ width: '13%' }} />
                        <col style={{ width: '12%' }} />
                    </colgroup>
                    <thead>
                        <tr style={{ background: '#f2f2f2' }}>
                            <th style={{ ...c, textAlign: 'center', fontSize: '6.5pt' }}>Indicate if Subject is<br />CORE, APPLIED, or SPECIALIZED</th>
                            <th style={{ ...c, textAlign: 'center' }}>SUBJECTS</th>
                            <th style={{ ...c, textAlign: 'center' }} colSpan="2">
                                Quarter
                                <div style={{ display: 'flex', borderTop: '1px solid #000', margin: '1px -3px -1px' }}>
                                    <div style={{ flex: 1, textAlign: 'center', borderRight: '1px solid #000', padding: '1px' }}>1ST</div>
                                    <div style={{ flex: 1, textAlign: 'center', padding: '1px' }}>2ND</div>
                                </div>
                            </th>
                            <th style={{ ...c, textAlign: 'center' }}>SEM FINAL<br />GRADE</th>
                            <th style={{ ...c, textAlign: 'center' }}>ACTION<br />TAKEN</th>
                        </tr>
                    </thead>
                    <tbody>
                        {Array.from({ length: SUBJECT_ROWS }).map((_, i) => {
                            const s = subjects[i] || {};
                            return (
                                <tr key={i}>
                                    <td style={{ ...c, textAlign: 'center', fontSize: '7pt' }}>{s.type}</td>
                                    <td style={c}>{s.subject}</td>
                                    <td style={{ ...c, textAlign: 'center' }}>{s.q1}</td>
                                    <td style={{ ...c, textAlign: 'center' }}>{s.q2}</td>
                                    <td style={{ ...c, textAlign: 'center', fontWeight: s.final ? 'bold' : 'normal' }}>{s.final}</td>
                                    <td style={{
                                        ...c,
                                        textAlign: 'center',
                                        fontWeight: 'bold',
                                        fontSize: '7pt',
                                        color: s.action === 'PASSED' ? '#006100' : s.action === 'FAILED' ? '#9c0006' : '#000',
                                    }}>{s.action}</td>
                                </tr>
                            );
                        })}
                    </tbody>
                </table>

                {/* Gen Ave + Remarks */}
                <table style={{ ...tbl, marginTop: '-1px' }}>
                    <tbody>
                        <tr>
                            <td style={{ ...cl, width: '75%', textAlign: 'right', paddingRight: '8px' }}>General Ave. for the Semester:</td>
                            <td style={{ ...cv, textAlign: 'center', fontSize: '9pt', fontWeight: 'bold' }}>{sem.genAve}</td>
                        </tr>
                        <tr>
                            <td style={{ ...c, padding: '2px 4px' }} colSpan="2">
                                <strong>REMARKS:</strong> {sem.remarks}
                            </td>
                        </tr>
                    </tbody>
                </table>

                {/* Signatories */}
                <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '8px', marginBottom: '4px' }}>
                    <tbody>
                        <tr>
                            <td style={{ width: '33%', textAlign: 'center', paddingTop: '16px', fontSize: '7pt' }}>
                                <div style={{ borderBottom: '1px solid #000', width: '85%', margin: '0 auto', fontWeight: 'bold', minHeight: '12px', paddingBottom: '2px' }}>{sem.adviserName}</div>
                                <div style={{ fontSize: '6pt', marginTop: '1px' }}>Signature of Adviser over Printed Name</div>
                            </td>
                            <td style={{ width: '34%', textAlign: 'center', paddingTop: '16px', fontSize: '7pt' }}>
                                <div style={{ borderBottom: '1px solid #000', width: '85%', margin: '0 auto', fontWeight: 'bold', minHeight: '12px', paddingBottom: '2px' }}>{sem.certName}</div>
                                <div style={{ fontSize: '6pt', marginTop: '1px' }}>Signature of Authorized Person over Printed Name, Designation</div>
                            </td>
                            <td style={{ width: '33%', textAlign: 'center', paddingTop: '16px', fontSize: '7pt' }}>
                                <div style={{ borderBottom: '1px solid #000', width: '85%', margin: '0 auto', minHeight: '12px', paddingBottom: '2px' }}>{sem.dateChecked}</div>
                                <div style={{ fontSize: '6pt', marginTop: '1px' }}>Date Checked (MM/DD/YYYY)</div>
                            </td>
                        </tr>
                    </tbody>
                </table>

                {/* Remedial Classes */}
                <div style={subHdr}>REMEDIAL CLASSES</div>
                <table style={tbl}>
                    <tbody>
                        <tr>
                            <td style={cl}>Conducted from (MM/DD/YYYY):</td>
                            <td style={cv}>{remedial.from}</td>
                            <td style={cl}>to (MM/DD/YYYY):</td>
                            <td style={cv}>{remedial.to}</td>
                            <td style={cl}>SCHOOL:</td>
                            <td style={cv}>{remedial.school}</td>
                            <td style={cl}>SCHOOL ID:</td>
                            <td style={cv}>{remedial.schoolId}</td>
                        </tr>
                    </tbody>
                </table>
                <table style={{ ...tbl, marginTop: '-1px' }}>
                    <colgroup>
                        <col style={{ width: '22%' }} />
                        <col style={{ width: '28%' }} />
                        <col style={{ width: '12%' }} />
                        <col style={{ width: '14%' }} />
                        <col style={{ width: '14%' }} />
                        <col style={{ width: '10%' }} />
                    </colgroup>
                    <thead>
                        <tr style={{ background: '#f2f2f2' }}>
                            <th style={{ ...c, textAlign: 'center', fontSize: '6.5pt' }}>Indicate if Subject is<br />CORE, APPLIED, or SPECIALIZED</th>
                            <th style={{ ...c, textAlign: 'center' }}>SUBJECTS</th>
                            <th style={{ ...c, textAlign: 'center', fontSize: '6.5pt' }}>SEM FINAL<br />GRADE</th>
                            <th style={{ ...c, textAlign: 'center', fontSize: '6.5pt' }}>REMEDIAL<br />CLASS MARK</th>
                            <th style={{ ...c, textAlign: 'center', fontSize: '6.5pt' }}>RECOMPUTED<br />FINAL GRADE</th>
                            <th style={{ ...c, textAlign: 'center', fontSize: '6.5pt' }}>ACTION<br />TAKEN</th>
                        </tr>
                    </thead>
                    <tbody>
                        {Array.from({ length: Math.max(3, remSubjects.length) }).map((_, i) => {
                            const rs = remSubjects[i] || {};
                            return (
                                <tr key={i}>
                                    <td style={{ ...c, textAlign: 'center', fontSize: '7pt' }}>{rs.type}</td>
                                    <td style={c}>{rs.subject}</td>
                                    <td style={{ ...c, textAlign: 'center' }}>{rs.semGrade}</td>
                                    <td style={{ ...c, textAlign: 'center' }}>{rs.remedialMark}</td>
                                    <td style={{ ...c, textAlign: 'center' }}>{rs.recomputedGrade}</td>
                                    <td style={{ ...c, textAlign: 'center' }}>{rs.action}</td>
                                </tr>
                            );
                        })}
                    </tbody>
                </table>
                <div style={{ fontSize: '7pt', marginTop: '4px' }}>
                    <strong>Name of Teacher/Adviser:</strong>{' '}
                    <span style={{ borderBottom: '1px solid #000', display: 'inline-block', minWidth: '140px' }}>{remedial.teacherName}</span>
                    <span style={{ float: 'right' }}>
                        <strong>Signature:</strong> <span style={{ borderBottom: '1px solid #000', display: 'inline-block', minWidth: '120px' }}></span>
                    </span>
                </div>
            </div>
        );
    };

    return (
        <div>
            {/* ═══════════════════════ PAGE 1 — FRONT ═══════════════════════ */}
            <div style={page}>
                {/* ── Header ── */}
                <div style={{ textAlign: 'center', marginBottom: '6px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                        <div style={{ width: '60px' }}></div>
                        <div style={{ textAlign: 'center' }}>
                            <div style={{ fontSize: '9pt' }}>REPUBLIC OF THE PHILIPPINES</div>
                            <div style={{ fontSize: '9pt', fontWeight: 'bold' }}>DEPARTMENT OF EDUCATION</div>
                        </div>
                        <div style={{ fontSize: '7pt', fontWeight: 'bold', textAlign: 'right' }}>FORM 137-SHS</div>
                    </div>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold', letterSpacing: '1px', margin: '2px 0' }}>
                        SENIOR HIGH SCHOOL STUDENT PERMANENT RECORD
                    </div>
                </div>

                {/* ── LEARNER'S INFORMATION ── */}
                <div style={hdr}>LEARNER'S INFORMATION</div>
                <table style={tbl}>
                    <colgroup>
                        <col style={{ width: '12%' }} />
                        <col style={{ width: '22%' }} />
                        <col style={{ width: '12%' }} />
                        <col style={{ width: '22%' }} />
                        <col style={{ width: '14%' }} />
                        <col style={{ width: '18%' }} />
                    </colgroup>
                    <tbody>
                        <tr>
                            <td style={cl}>LAST NAME:</td>
                            <td style={cv}>{info.lname}</td>
                            <td style={cl}>FIRST NAME:</td>
                            <td style={cv}>{info.fname}</td>
                            <td style={cl}>MIDDLE NAME:</td>
                            <td style={cv}>{info.mname}</td>
                        </tr>
                        <tr>
                            <td style={cl}>LRN:</td>
                            <td style={cv}>{info.lrn}</td>
                            <td style={{ ...cl, fontSize: '6.5pt' }}>Date of Birth (MM/DD/YYYY):</td>
                            <td style={cv}>{info.birthdate}</td>
                            <td style={cl}>Sex:</td>
                            <td style={cv}>{info.sex}</td>
                        </tr>
                        <tr>
                            <td style={{ ...cl, fontSize: '6.5pt' }} colSpan="2">Date of SHS Admission (MM/DD/YYYY):</td>
                            <td style={cv} colSpan="4">{info.admissionDate}</td>
                        </tr>
                    </tbody>
                </table>

                {/* ── ELIGIBILITY ── */}
                <div style={{ ...hdr, marginTop: '4px' }}>ELIGIBILITY FOR SHS ENROLMENT</div>
                <table style={tbl}>
                    <tbody>
                        <tr>
                            <td style={c} width="20%">
                                <input type="checkbox" checked={!!elig.hsCompleter} readOnly style={{ transform: 'scale(0.75)', marginRight: '2px', verticalAlign: 'middle' }} />
                                <span style={{ fontSize: '7pt' }}>High School Completer*</span>
                            </td>
                            <td style={cl} width="8%">Gen. Ave:</td>
                            <td style={cv} width="10%">{elig.hsGenAve}</td>
                            <td style={c} width="22%">
                                <input type="checkbox" checked={!!elig.jhsCompleter} readOnly style={{ transform: 'scale(0.75)', marginRight: '2px', verticalAlign: 'middle' }} />
                                <span style={{ fontSize: '7pt' }}>Junior High School Completer</span>
                            </td>
                            <td style={cl} width="8%">Gen. Ave:</td>
                            <td style={cv} width="10%">{elig.jhsGenAve}</td>
                        </tr>
                        <tr>
                            <td style={{ ...cl, fontSize: '6.5pt' }} colSpan="2">Date of Graduation/Completion (MM/DD/YYYY):</td>
                            <td style={cv}>{elig.gradDate}</td>
                            <td style={cl}>Name of School:</td>
                            <td style={cv} colSpan="2">{elig.schoolName}</td>
                        </tr>
                        <tr>
                            <td style={cl}>School Address:</td>
                            <td style={cv} colSpan="5">{elig.schoolAddress}</td>
                        </tr>
                        <tr>
                            <td style={c}>
                                <input type="checkbox" checked={!!elig.pept} readOnly style={{ transform: 'scale(0.75)', marginRight: '2px', verticalAlign: 'middle' }} />
                                <span style={{ fontSize: '7pt' }}>PEPT Passer**</span>
                            </td>
                            <td style={cl}>Rating:</td>
                            <td style={cv}>{elig.peptRating}</td>
                            <td style={c}>
                                <input type="checkbox" checked={!!elig.als} readOnly style={{ transform: 'scale(0.75)', marginRight: '2px', verticalAlign: 'middle' }} />
                                <span style={{ fontSize: '7pt' }}>ALS A&amp;E Passer***</span>
                            </td>
                            <td style={cl}>Rating:</td>
                            <td style={cv}>{elig.alsRating}</td>
                        </tr>
                        <tr>
                            <td style={{ ...cl, fontSize: '6.5pt' }} colSpan="2">Date of Examination/Assessment (MM/DD/YYYY):</td>
                            <td style={cv}>{elig.examDate}</td>
                            <td style={{ ...cl, fontSize: '6.5pt' }} colSpan="2">Name and Address of Community Learning Center:</td>
                            <td style={cv}>{elig.clcName}</td>
                        </tr>
                        <tr>
                            <td style={c} colSpan="2">
                                <input type="checkbox" checked={!!elig.others} readOnly style={{ transform: 'scale(0.75)', marginRight: '2px', verticalAlign: 'middle' }} />
                                <span style={{ fontSize: '7pt' }}>Others (Pls. Specify):</span>
                            </td>
                            <td style={cv} colSpan="4">{elig.othersSpec}</td>
                        </tr>
                    </tbody>
                </table>
                <div style={{ fontSize: '6pt', marginTop: '2px', lineHeight: 1.3, color: '#333' }}>
                    <div>*High School Completers are students who graduated from secondary school under the old curriculum</div>
                    <div>**PEPT — Philippine Educational Placement Test for JHS</div>
                    <div>***ALS A&amp;E — Alternative Learning System Accreditation and Equivalency Test for JHS</div>
                </div>

                {/* ── Semester Block 1 (Grade 11 Sem 1) ── */}
                <div style={{ marginTop: '8px' }}>
                    {renderSemBlock(sem1, 'Semester 1')}
                </div>

                {/* ── Semester Block 2 (Grade 11 Sem 2) ── */}
                <div style={{ marginTop: '8px' }}>
                    {renderSemBlock(sem2, 'Semester 2')}
                </div>
            </div>

            {/* ═══════════════════════ PAGE 2 — BACK ═══════════════════════ */}
            <div style={page}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px' }}>
                    <div style={{ fontSize: '7pt' }}>Page 2</div>
                    <div style={{ fontSize: '7pt', fontWeight: 'bold' }}>Form 137-SHS</div>
                </div>

                {/* Note: In the actual 4-semester template, blocks 3 & 4 go here.
                    Since data model currently has semester1/semester2 only,
                    we show empty blocks for Grade 12 semesters. */}

                {/* ── CERTIFICATION ── */}
                <div style={{ marginTop: '12px', borderTop: '2px solid #000', paddingTop: '8px' }}>
                    <table style={tbl}>
                        <tbody>
                            <tr>
                                <td style={cl} width="20%">Track/Strand Accomplished:</td>
                                <td style={cv} width="40%">{cert.trackStrand}</td>
                                <td style={cl} width="18%">SHS General Average:</td>
                                <td style={cv} width="22%">{cert.genAve}</td>
                            </tr>
                            <tr>
                                <td style={cl}>Awards/Honors Received:</td>
                                <td style={cv}>{cert.awards}</td>
                                <td style={{ ...cl, fontSize: '6.5pt' }}>Date of SHS Graduation (MM/DD/YYYY):</td>
                                <td style={cv}>{cert.gradDate}</td>
                            </tr>
                        </tbody>
                    </table>

                    {/* Certified by */}
                    <div style={{ marginTop: '8px', fontSize: '7pt' }}>
                        <strong>Certified by:</strong>
                        <span style={{ marginLeft: '120px' }}><strong>Place School Seal Here:</strong></span>
                    </div>
                    <table style={{ width: '100%', marginTop: '20px' }}>
                        <tbody>
                            <tr>
                                <td style={{ width: '40%', textAlign: 'center', fontSize: '7pt' }}>
                                    <div style={{ borderBottom: '1px solid #000', width: '70%', margin: '0 auto', fontWeight: 'bold', minHeight: '14px' }}>{cert.schoolHead}</div>
                                    <div style={{ fontSize: '6pt' }}>Signature of School Head over Printed Name</div>
                                </td>
                                <td style={{ width: '20%', textAlign: 'center', fontSize: '7pt' }}>
                                    <div style={{ borderBottom: '1px solid #000', width: '70%', margin: '0 auto', minHeight: '14px' }}>{cert.certDate}</div>
                                    <div style={{ fontSize: '6pt' }}>Date</div>
                                </td>
                                <td style={{ width: '40%' }}></td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                {/* NOTE */}
                <div style={{ marginTop: '16px', border: '1px solid #999', padding: '6px 8px', fontSize: '6.5pt', lineHeight: 1.4, color: '#333' }}>
                    <strong>NOTE:</strong><br />
                    This permanent record or a photocopy of this permanent record that bears the seal of the school and the original signature in ink of the School Head shall be considered valid for all legal purposes. Any erasure or alteration made on this copy should be validated by the School Head.<br />
                    If the student transfers to another school, the originating school should produce one (1) certified true copy of this permanent record for safekeeping. The receiving school shall continue filling up the original form.<br />
                    Upon graduation, the school from which the student graduated should keep the original form and produce one (1) certified true copy for the Division Office.
                </div>

                {/* REMARKS */}
                <div style={{ marginTop: '10px', border: '1px solid #000', padding: '6px 8px', minHeight: '40px', fontSize: '7pt' }}>
                    <strong>REMARKS:</strong> (Please indicate the purpose for which this permanent record will be used)<br />
                    {cert.remarks}
                </div>

                {/* Date Issued */}
                <div style={{ marginTop: '8px', fontSize: '7pt' }}>
                    <strong>Date Issued (MM/DD/YYYY):</strong>{' '}
                    <span style={{ borderBottom: '1px solid #000', display: 'inline-block', minWidth: '120px', paddingBottom: '2px' }}>{cert.dateIssued}</span>
                </div>
            </div>
        </div>
    );
}
