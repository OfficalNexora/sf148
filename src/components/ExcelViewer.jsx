import React from 'react';
import '../styles/excel-preview.css';

const ExcelViewer = ({ studentData, onBack }) => {
    if (!studentData) return null;

    const { info } = studentData;

    const renderSemester = (semKey, title) => {
        const sem = studentData[semKey];
        if (!sem || sem.subjects.length === 0) return null;

        return (
            <div className="excel-sheet-preview">
                <table className="no-border">
                    <tbody>
                        <tr>
                            <td width="10%">SCHOOL:</td>
                            <td className="excel-value" width="40%">{sem.school}</td>
                            <td width="12%">SCHOOL ID:</td>
                            <td className="excel-value" width="10%">{sem.schoolId}</td>
                            <td width="13%">GRADE LEVEL:</td>
                            <td className="excel-value" width="5%">{sem.gradeLevel}</td>
                            <td width="5%">SY:</td>
                            <td className="excel-value" width="10%">{sem.sy}</td>
                            <td width="5%">SEM:</td>
                            <td className="excel-value" width="5%">{sem.sem}</td>
                        </tr>
                        <tr>
                            <td>TRACK/STRAND:</td>
                            <td className="excel-value" colSpan={3}>{sem.trackStrand}</td>
                            <td>SECTION:</td>
                            <td className="excel-value" colSpan={5}>{sem.section}</td>
                        </tr>
                    </tbody>
                </table>

                <table className="subjects-table">
                    <colgroup>
                        <col style={{ width: '12%' }} />
                        <col style={{ width: '40%' }} />
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '8%' }} />
                        <col style={{ width: '16%' }} />
                        <col style={{ width: '16%' }} />
                    </colgroup>
                    <thead>
                        <tr>
                            <th>TYPE</th>
                            <th>SUBJECTS</th>
                            <th>Q1</th>
                            <th>Q2</th>
                            <th>SEM FINAL GRADE</th>
                            <th>ACTION TAKEN</th>
                        </tr>
                    </thead>
                    <tbody>
                        {sem.subjects.map((s, i) => (
                            <tr key={i}>
                                <td>{s.type}</td>
                                <td className="subject-name">{s.subject}</td>
                                <td>{s.q1}</td>
                                <td>{s.q2}</td>
                                <td>{s.final}</td>
                                <td style={{ color: s.action === 'FAILED' ? 'red' : 'inherit' }}>{s.action}</td>
                            </tr>
                        ))}
                        {/* Fill empty rows to maintain height if needed */}
                        {[...Array(Math.max(0, 8 - sem.subjects.length))].map((_, i) => (
                            <tr key={`empty-${i}`}>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                            </tr>
                        ))}
                    </tbody>
                    <tfoot>
                        <tr>
                            <td colSpan={2} style={{ border: 'none' }}>&nbsp;</td>
                            <td colSpan={2} style={{ textAlign: 'right', fontWeight: 'bold', border: 'none', fontSize: '9px' }}>General Ave. for the Semester:</td>
                            <td style={{ fontWeight: 'bold', fontSize: '11px' }}>{sem.genAve}</td>
                            <td className="excel-value" style={{ border: 'none', textAlign: 'left', paddingLeft: '10px' }}>{sem.remarks}</td>
                        </tr>
                    </tfoot>
                </table>

                <table className="no-border" style={{ marginTop: '5px' }}>
                    <tbody>
                        <tr>
                            <td width="15%">Prepared by:</td>
                            <td className="excel-value" width="30%">{sem.adviserName}</td>
                            <td width="20%">Certified True & Correct:</td>
                            <td className="excel-value" width="35%">{sem.certName}</td>
                        </tr>
                        <tr>
                            <td className="excel-label">Signature of Adviser over Printed Name</td>
                            <td></td>
                            <td className="excel-label">Signature of Authorized Person over Printed Name, Designation</td>
                            <td></td>
                        </tr>
                    </tbody>
                </table>

                {renderRemedial(semKey)}
                <div style={{ height: '20px' }}></div>
            </div>
        );
    };

    const renderRemedial = (semKey) => {
        const sem = studentData[semKey];
        if (!sem || !sem.remedial || sem.remedial.subjects.length === 0) return null;

        const rem = sem.remedial;

        return (
            <div className="remedial-excel" style={{ marginTop: '10px' }}>
                <div className="excel-section-title">REMEDIAL CLASSES</div>
                <table className="no-border">
                    <tbody>
                        <tr>
                            <td width="15%">Conducted from:</td>
                            <td className="excel-value" width="15%">{rem.from}</td>
                            <td width="5%">to:</td>
                            <td className="excel-value" width="15%">{rem.to}</td>
                            <td width="10%">SCHOOL:</td>
                            <td className="excel-value" width="25%">{rem.school}</td>
                            <td width="10%">SCHOOL ID:</td>
                            <td className="excel-value" width="5%">{rem.schoolId}</td>
                        </tr>
                    </tbody>
                </table>
                <table className="subjects-table">
                    <colgroup>
                        <col style={{ width: '12%' }} />
                        <col style={{ width: '38%' }} />
                        <col style={{ width: '12.5%' }} />
                        <col style={{ width: '12.5%' }} />
                        <col style={{ width: '12.5%' }} />
                        <col style={{ width: '12.5%' }} />
                    </colgroup>
                    <thead>
                        <tr>
                            <th>TYPE</th>
                            <th>SUBJECTS</th>
                            <th>SEM FINAL GRADE</th>
                            <th>REMEDIAL MARK</th>
                            <th>RECOMPUTED FINAL</th>
                            <th>ACTION TAKEN</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rem.subjects.map((s, i) => (
                            <tr key={i}>
                                <td>{s.type}</td>
                                <td className="subject-name">{s.subject}</td>
                                <td>{s.final}</td>
                                <td>{s.remedialMark}</td>
                                <td>{s.recomputedFinal}</td>
                                <td>{s.action}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        );
    };

    return (
        <div className="excel-preview-container">
            <button className="btn-secondary" onClick={onBack} style={{ position: 'absolute', top: '20px', left: '20px' }}>
                ← Back to Editor
            </button>
            <button className="btn-print" onClick={() => window.print()}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 6 2 18 2 18 9"></polyline><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect></svg>
                Print SF10
            </button>

            {/* Page 1: FRONT */}
            <div className="excel-page">
                <div className="excel-header-text">
                    REPUBLIC OF THE PHILIPPINES<br />
                    DEPARTMENT OF EDUCATION<br />
                    <span style={{ fontSize: '14px' }}>SENIOR HIGH SCHOOL STUDENT PERMANENT RECORD</span>
                </div>
                <div style={{ textAlign: 'right', fontSize: '12px', fontWeight: 'bold' }}>SF10-SHS</div>

                <div className="excel-section-title" style={{ marginTop: '10px' }}>LEARNER'S INFORMATION</div>
                <table className="no-border">
                    <tbody>
                        <tr>
                            <td width="10%">LAST NAME:</td>
                            <td className="excel-value" width="25%">{info.lname}</td>
                            <td width="10%">FIRST NAME:</td>
                            <td className="excel-value" width="25%">{info.fname}</td>
                            <td width="12%">MIDDLE NAME:</td>
                            <td className="excel-value" width="18%">{info.mname}</td>
                        </tr>
                        <tr>
                            <td>LRN:</td>
                            <td className="excel-value">{info.lrn}</td>
                            <td>Date of Birth:</td>
                            <td className="excel-value">{info.birthdate}</td>
                            <td>Sex:</td>
                            <td className="excel-value">{info.sex}</td>
                        </tr>
                    </tbody>
                </table>

                <div className="excel-section-title" style={{ marginTop: '10px' }}>SCHOLASTIC RECORD</div>
                {renderSemester('semester1', 'Grade 11 - 1st Sem')}
                {renderSemester('semester2', 'Grade 11 - 2nd Sem')}
            </div>

            {/* Page 2: BACK */}
            <div className="excel-page">
                <div style={{ textAlign: 'right', fontSize: '10px' }}>Page 2</div>
                {renderSemester('semester3', 'Grade 12 - 1st Sem')}
                {renderSemester('semester4', 'Grade 12 - 2nd Sem')}

                <div className="excel-section-title" style={{ marginTop: '10px' }}>SUMMARY</div>
                <table className="no-border">
                    <tbody>
                        <tr>
                            <td width="25%">Track/Strand Accomplished:</td>
                            <td className="excel-value">{info.trackStrandAccomplished || '—'}</td>
                            <td width="20%">SHS General Average:</td>
                            <td className="excel-value" width="10%">{info.shsGenAve || '—'}</td>
                        </tr>
                        <tr>
                            <td>Awards/Honors Received:</td>
                            <td className="excel-value" colSpan={3}>{info.awards || '—'}</td>
                        </tr>
                        <tr>
                            <td>Date of SHS Graduation:</td>
                            <td className="excel-value">{info.gradDate || '—'}</td>
                        </tr>
                    </tbody>
                </table>

                <div style={{ marginTop: '40px', display: 'flex', justifyContent: 'space-between' }}>
                    <div style={{ width: '45%', textAlign: 'center' }}>
                        <div style={{ borderBottom: '1px solid black', paddingBottom: '2px' }}>&nbsp;</div>
                        <div className="excel-label">Signature of School Head over Printed Name</div>
                    </div>
                    <div style={{ width: '20%', border: '1px solid #ccc', height: '100px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '10px', color: '#888' }}>
                        Place School Seal Here
                    </div>
                    <div style={{ width: '25%', textAlign: 'center' }}>
                        <div style={{ borderBottom: '1px solid black', paddingBottom: '2px' }}>&nbsp;</div>
                        <div className="excel-label">Date</div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default ExcelViewer;
