import React, { useState, useCallback, useRef } from 'react';
import { read, utils } from 'xlsx';
// db service available if needed in future
import PlaceholderChecker from './PlaceholderChecker';

// --- Icons ---
const UserIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
);
const CheckIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
);
const BookIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>
);
const GradIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 10v6M2 10l10-5 10 5-10 5z"></path><path d="M6 12v5c3 3 9 3 12 0v-5"></path></svg>
);
const PrintIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 6 2 18 2 18 9"></polyline><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect></svg>
);
const SearchIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
);

const TABS = [
    { key: 'info', label: 'Learner Info', icon: <UserIcon /> },
    { key: 'eligibility', label: 'Eligibility', icon: <CheckIcon /> },
    { key: 'front', label: 'Front Page (Grade 11)', icon: <BookIcon /> },
    { key: 'back', label: 'Back Page (Grade 12)', icon: <BookIcon /> },
    { key: 'annex', label: 'Annex (Master List)', icon: <SearchIcon /> },
    { key: 'certification', label: 'Certification', icon: <GradIcon /> }
];

const EMPTY_SUBJECT = { type: '', subject: '', q1: '', q2: '', final: '', action: '' };
const EMPTY_REMEDIAL_SUBJECT = { type: '', subject: '', semGrade: '', remedialMark: '', recomputedGrade: '', action: '' };

const STRAND_SUBJECTS = {
    'STEM': [
        { type: 'Core', subject: 'Oral Communication in Context' },
        { type: 'Core', subject: 'Reading and Writing Skills' },
        { type: 'Core', subject: 'Komunikasyon at Pananaliksik sa Wika at Kulturang Pilipino' },
        { type: 'Core', subject: 'Pagbasa at Pagsusuri ng Iba’t-ibang Teksto Tungo sa Pananaliksik' },
        { type: 'Core', subject: '21st Century Literature from the Philippines and the World' },
        { type: 'Core', subject: 'Contemporary Philippine Arts from the Regions' },
        { type: 'Core', subject: 'Media and Information Literacy' },
        { type: 'Core', subject: 'General Mathematics' },
        { type: 'Core', subject: 'Statistics and Probability' },
        { type: 'Core', subject: 'Earth and Life Science' },
        { type: 'Core', subject: 'Physical Science' },
        { type: 'Core', subject: 'Personal Development/Pansariling Kaunlaran' },
        { type: 'Core', subject: 'Understanding Culture, Society and Politics' },
        { type: 'Core', subject: 'Physical Education and Health' },
        { type: 'Applied', subject: 'English for Academic and Professional Purposes' },
        { type: 'Applied', subject: 'Practical Research 1' },
        { type: 'Applied', subject: 'Practical Research 2' },
        { type: 'Applied', subject: 'Pagsulat sa Filipino sa Piling Larangan' },
        { type: 'Applied', subject: 'Empowerment Technologies' },
        { type: 'Applied', subject: 'Entrepreneurship' },
        { type: 'Specialized', subject: 'Pre-Calculus' },
        { type: 'Specialized', subject: 'Basic Calculus' },
        { type: 'Specialized', subject: 'General Biology 1' },
        { type: 'Specialized', subject: 'General Biology 2' },
        { type: 'Specialized', subject: 'General Physics 1' },
        { type: 'Specialized', subject: 'General Physics 2' },
        { type: 'Specialized', subject: 'General Chemistry 1' },
        { type: 'Specialized', subject: 'General Chemistry 2' }
    ],
    'HUMSS': [
        { type: 'Core', subject: 'Oral Communication in Context' },
        { type: 'Core', subject: 'Reading and Writing Skills' },
        { type: 'Core', subject: 'General Mathematics' },
        { type: 'Core', subject: 'Statistics and Probability' },
        { type: 'Specialized', subject: 'Creative Writing' },
        { type: 'Specialized', subject: 'Introduction to World Religions and Belief Systems' },
        { type: 'Specialized', subject: 'Creative Nonfiction' },
        { type: 'Specialized', subject: 'Trends, Networks, and Critical Thinking in the 21st Century Culture' }
    ],
    'GAS': [
        { type: 'Core', subject: 'Oral Communication in Context' },
        { type: 'Core', subject: 'Reading and Writing Skills' },
        { type: 'Core', subject: 'General Mathematics' },
        { type: 'Core', subject: 'Statistics and Probability' },
        { type: 'Applied', subject: 'Practical Research 1' },
        { type: 'Applied', subject: 'Practical Research 2' }
    ],
    'TVL': [
        { type: 'Core', subject: 'Oral Communication in Context' },
        { type: 'Core', subject: 'Reading and Writing Skills' },
        { type: 'Core', subject: 'General Mathematics' },
        { type: 'Applied', subject: 'Empowerment Technologies' }
    ]
};

// ── Field helper (Defined outside to prevent re-mounting and focus loss) ──
const Field = ({ label, value, onInput, type = 'text', placeholder = '', fullSpan = false, ...rest }) => (
    <div className={`field-group${fullSpan ? ' full-span' : ''}`}>
        <label>{label}</label>
        {type === 'textarea' ? (
            <textarea value={value || ''} onChange={(e) => onInput(e.target.value)} placeholder={placeholder} {...rest} />
        ) : (
            <input type={type} value={value || ''} onChange={(e) => onInput(e.target.value)} placeholder={placeholder} {...rest} />
        )}
    </div>
);

// K-12 standard subjects from the official PLACEHOLDER(ALL).xlsx template
// Hardcoded for instant load — no runtime fetch needed
const K12_TEMPLATE_SUBJECTS = {
    Core: [
        '21st Century Literature from the Philippines and the World',
        'Contemporary Philippine Arts from the Regions',
        'Disaster Readiness and Risk Reduction',
        'Earth and Life Science*',
        'Earth Science',
        'General Mathematics',
        'Introduction to the Philosophy of the Human Person/Pambungad sa Pilosopiya ng Tao',
        'Komunikasyon at Pananaliksik sa Wika at Kulturang Pilipino',
        'Media and Information Literacy',
        'Oral Communication',
        'Pagbasa at Pagsusuri ng Iba\'t Ibang Teksto Tungo sa Pananaliksik',
        'Personal Development/Pansariling Kaunlaran',
        'Physical Education and Health',
        'Physical Science*',
        'Reading and Writing',
        'Statistics and Probability',
        'Understanding Culture, Society and Politics',
    ],
    Applied: [
        'Empowerment Technologies',
        'English for Academic and Professional Purposes',
        'Entrepreneurship',
        'Filipino sa Piling Larang',
        'Inquiries, Investigations and Immersion',
        'Practical Research 1',
        'Practical Research 2',
    ],
    Specialized: [
        'Computer Programming (Oracle) NC III',
    ],
    Other: [],
};

function Form137Preview({ data }) {
    const fullName = [data.info.lname, data.info.fname, data.info.mname].filter(Boolean).join(', ');
    const semesterKeys = ['semester1', 'semester2', 'semester3', 'semester4'];

    return (
        <div style={{ color: '#111827', fontFamily: 'Arial, sans-serif', fontSize: '12px' }}>
            <h2 style={{ marginTop: 0, marginBottom: '10px' }}>Form 137 Snapshot</h2>
            <div style={{ marginBottom: '14px', lineHeight: 1.5 }}>
                <div><strong>Name:</strong> {fullName || '-'}</div>
                <div><strong>LRN:</strong> {data.info.lrn || '-'}</div>
                <div><strong>Sex:</strong> {data.info.sex || '-'}</div>
                <div><strong>Birthdate:</strong> {data.info.birthdate || '-'}</div>
            </div>

            {semesterKeys.map((semKey, index) => {
                const sem = data[semKey];
                if (!sem) return null;
                const subjects = (sem.subjects || []).filter(s => s.subject);

                return (
                    <div key={semKey} style={{ marginBottom: '16px' }}>
                        <h3 style={{ margin: '0 0 8px 0' }}>Semester {index + 1}</h3>
                        <div style={{ marginBottom: '6px' }}>
                            <strong>School:</strong> {sem.school || '-'} | <strong>SY:</strong> {sem.sy || '-'} | <strong>Section:</strong> {sem.section || '-'}
                        </div>
                        <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #d1d5db' }}>
                            <thead>
                                <tr style={{ background: '#f3f4f6' }}>
                                    <th style={{ border: '1px solid #d1d5db', padding: '4px' }}>Type</th>
                                    <th style={{ border: '1px solid #d1d5db', padding: '4px' }}>Subject</th>
                                    <th style={{ border: '1px solid #d1d5db', padding: '4px' }}>Q1</th>
                                    <th style={{ border: '1px solid #d1d5db', padding: '4px' }}>Q2</th>
                                    <th style={{ border: '1px solid #d1d5db', padding: '4px' }}>Final</th>
                                    <th style={{ border: '1px solid #d1d5db', padding: '4px' }}>Action</th>
                                </tr>
                            </thead>
                            <tbody>
                                {subjects.length === 0 ? (
                                    <tr>
                                        <td colSpan={6} style={{ border: '1px solid #d1d5db', padding: '6px', textAlign: 'center' }}>
                                            No subjects yet.
                                        </td>
                                    </tr>
                                ) : (
                                    subjects.map((subj, i) => (
                                        <tr key={`${semKey}-${i}`}>
                                            <td style={{ border: '1px solid #d1d5db', padding: '4px' }}>{subj.type || ''}</td>
                                            <td style={{ border: '1px solid #d1d5db', padding: '4px' }}>{subj.subject || ''}</td>
                                            <td style={{ border: '1px solid #d1d5db', padding: '4px' }}>{subj.q1 || ''}</td>
                                            <td style={{ border: '1px solid #d1d5db', padding: '4px' }}>{subj.q2 || ''}</td>
                                            <td style={{ border: '1px solid #d1d5db', padding: '4px' }}>{subj.final || ''}</td>
                                            <td style={{ border: '1px solid #d1d5db', padding: '4px' }}>{subj.action || ''}</td>
                                        </tr>
                                    ))
                                )}
                            </tbody>
                        </table>
                    </div>
                );
            })}
        </div>
    );
}

function StudentEditor({ data, onChange, onSave, isDesktopMode = false }) {
    const [activeTab, setActiveTab] = useState('info');
    const [isPrinting, setIsPrinting] = useState(false);
    const [showChecker, setShowChecker] = useState(false);
    const [showRawData, setShowRawData] = useState(false);
    const saveTimerRef = useRef(null);
    const printActionLabel = 'Print to Excel';

    // Cleanup save timer on unmount
    React.useEffect(() => {
        return () => {
            if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
        };
    }, []);

    const scheduleAutoSave = useCallback((newData) => {
        if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
        saveTimerRef.current = setTimeout(() => onSave(newData), 1500);
    }, [onSave]);

    // ── Generic updaters ──

    const update = useCallback((section, field, value) => {
        const newData = { ...data, [section]: { ...data[section], [field]: value } };
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const updateSem = useCallback((semKey, field, value) => {
        const newData = { ...data, [semKey]: { ...data[semKey], [field]: value } };
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const updateSub = useCallback((semKey, idx, field, value) => {
        if (!data[semKey] || !data[semKey].subjects[idx]) return;

        // Performance optimization: don't update if value hasn't changed
        if (data[semKey].subjects[idx][field] === value) return;

        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].subjects = [...newData[semKey].subjects];

        // Basic Validation: Clamping grades 0-100
        let processedValue = value;
        if (['q1', 'q2', 'final'].includes(field) && value !== '') {
            const num = parseInt(value, 10);
            if (!isNaN(num)) {
                if (num > 100) processedValue = '100';
                if (num < 0) processedValue = '0';
            }
        }

        newData[semKey].subjects[idx] = { ...newData[semKey].subjects[idx], [field]: processedValue };
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const updateRem = useCallback((semKey, field, value) => {
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].remedial = { ...newData[semKey].remedial, [field]: value };
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const updateRemSub = useCallback((semKey, idx, field, value) => {
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].remedial = { ...newData[semKey].remedial };
        newData[semKey].remedial.subjects = [...newData[semKey].remedial.subjects];
        newData[semKey].remedial.subjects[idx] = { ...newData[semKey].remedial.subjects[idx], [field]: value };
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    // ── Add / Remove subject rows ──

    const addSubject = useCallback((semKey) => {
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].subjects = [...newData[semKey].subjects, { ...EMPTY_SUBJECT }];
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const removeSubject = useCallback((semKey) => {
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        if (newData[semKey].subjects.length <= 1) return;
        newData[semKey].subjects = newData[semKey].subjects.slice(0, -1);
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const addRemedialSubject = useCallback((semKey) => {
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].remedial = { ...newData[semKey].remedial };
        newData[semKey].remedial.subjects = [...newData[semKey].remedial.subjects, { ...EMPTY_REMEDIAL_SUBJECT }];
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const removeRemedialSubject = useCallback((semKey) => {
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].remedial = { ...newData[semKey].remedial };
        if (newData[semKey].remedial.subjects.length <= 1) return;
        newData[semKey].remedial.subjects = newData[semKey].remedial.subjects.slice(0, -1);
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    // ── Utility Helpers ──

    const loadTemplateSubjects = useCallback(() => {
        if ((data.annex || []).some(a => a.subject && a.subject.trim() !== '')) {
            if (!window.confirm('This will overwrite your current Annex Master List with the default K-12 subjects. Continue?')) return;
        }

        const core = K12_TEMPLATE_SUBJECTS.Core || [];
        const applied = K12_TEMPLATE_SUBJECTS.Applied || [];
        const specialized = K12_TEMPLATE_SUBJECTS.Specialized || [];
        const other = K12_TEMPLATE_SUBJECTS.Other || [];

        const newAnnex = Array(36).fill(null).map((_, i) => {
            let type = 'Other';
            if (i < 15) type = 'Core';
            else if (i < 22) type = 'Applied';
            else if (i < 31) type = 'Specialized';
            return { type, subject: '', active: true };
        });

        core.forEach((s, i) => { if (i < 15) newAnnex[i] = { ...newAnnex[i], subject: s }; });
        applied.forEach((s, i) => { if (i < 7) newAnnex[15 + i] = { ...newAnnex[15 + i], subject: s }; });
        specialized.forEach((s, i) => { if (i < 9) newAnnex[22 + i] = { ...newAnnex[22 + i], subject: s }; });
        other.forEach((s, i) => { if (i < 5) newAnnex[31 + i] = { ...newAnnex[31 + i], subject: s }; });

        const newData = { ...data, annex: newAnnex };
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const clearGrades = useCallback((semKey) => {
        if (!window.confirm('Are you sure you want to clear ALL grades for this semester? Subjects will remain.')) return;
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].subjects = newData[semKey].subjects.map(s => ({
            ...s, q1: '', q2: '', final: '', action: ''
        }));
        newData[semKey].genAve = '';
        newData[semKey].remarks = '';
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const copyFromPrevious = useCallback((semKey) => {
        const semNum = parseInt(semKey.replace('semester', ''));
        if (semNum <= 1) return;
        const prevSemKey = `semester${semNum - 1}`;
        const prevSem = data[prevSemKey];

        if (!prevSem || prevSem.subjects.length === 0) {
            alert('No subjects found in the previous semester to copy from.');
            return;
        }

        if (data[semKey].subjects.length > 0 && data[semKey].subjects[0].subject !== '') {
            if (!window.confirm('This will append subjects from the previous semester. Continue?')) return;
        }

        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        const copiedSubjects = prevSem.subjects.map(s => ({
            ...EMPTY_SUBJECT,
            type: s.type,
            subject: s.subject
        }));

        // Filter out empty rows if they exist
        newData[semKey].subjects = [
            ...newData[semKey].subjects.filter(s => s.subject !== ''),
            ...copiedSubjects
        ];

        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const autoPopulateFromStrand = useCallback((semKey) => {
        const strand = data[semKey].trackStrand?.toUpperCase() || '';
        let matchedKey = null;

        // Try to find a match in the dictionary
        Object.keys(STRAND_SUBJECTS).forEach(key => {
            if (strand.includes(key)) matchedKey = key;
        });

        if (!matchedKey) {
            alert(`Could not find a curriculum template matching "${strand}".\n\nTry entering a standard strand like STEM, HUMSS, GAS, or TVL in the field above.`);
            return;
        }

        if (data[semKey].subjects.length > 0 && data[semKey].subjects[0].subject !== '') {
            if (!window.confirm(`This will add the standard ${matchedKey} subjects to your list. Continue?`)) return;
        }

        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        const templateSubjs = STRAND_SUBJECTS[matchedKey].map(s => ({
            ...EMPTY_SUBJECT,
            ...s
        }));

        // Filter out empty rows, then append
        newData[semKey].subjects = [
            ...newData[semKey].subjects.filter(s => s.subject !== ''),
            ...templateSubjs
        ];

        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    const clearAllSubjects = useCallback((semKey) => {
        if (!window.confirm('Delete ALL subjects for this semester? This cannot be undone.')) return;
        const newData = { ...data };
        newData[semKey] = { ...newData[semKey] };
        newData[semKey].subjects = [{ ...EMPTY_SUBJECT }];
        newData[semKey].genAve = '';
        newData[semKey].remarks = '';
        onChange(newData);
        scheduleAutoSave(newData);
    }, [data, onChange, scheduleAutoSave]);

    // ── Handlers for Print ──
    const handlePrint = async () => {
        setIsPrinting(true);
        setTimeout(async () => {
            try {
                // Safe check for Electron environment
                const renderer = window.ipcRenderer || (window.electron && window.electron.ipcRenderer);

                if (renderer) {
                    // Desktop App Environment (Supports Auto-Open)
                    const result = await renderer.invoke('print-excel-form', data);
                    if (!result.success) {
                        alert('Failed to generate Excel file: ' + result.error);
                    }
                } else {
                    // Web Browser Environment (Vercel):
                    // 1) try local bridge server (desktop auto-open), 2) fallback to browser download.
                    const { openExcelViaBridge } = await import('../services/excelBridgeClient');
                    const bridgeResult = await openExcelViaBridge(data, {
                        autoPrint: true,
                        openAfterPrint: true
                    });
                    if (bridgeResult.success && !bridgeResult.warning) {
                        // Fully successful bridge print
                        setIsPrinting(false);
                        return;
                    }

                    if (bridgeResult.warning) {
                        console.warn('Bridge returned warning, falling back to browser download:', bridgeResult.warning);
                    }

                    // Fallback to purely generating and downloading the Excel file in the browser
                    const { generateExcelForm } = await import('../utils/excelGenerator');
                    const result = await generateExcelForm(data);
                    if (result.success) {
                        alert('Excel downloaded successfully!\n\nAs you do not have a local Excel app configured, please upload the downloaded file to Google Sheets, Microsoft 365, or another online Excel viewer to open and print it.');
                    } else {
                        alert(`Failed to generate Excel file.\nError: ${result.error}`);
                    }
                }
            } catch (e) {
                alert('Error printing: ' + e.message);
            }
            setIsPrinting(false);
        }, 50);
    };


    // ── Render Learner Info tab ──
    const renderInfo = () => (
        <div className="editor-section">
            <h3><span className="section-icon"><UserIcon /></span> Learner's Information</h3>
            <div className="form-grid three-col">
                <Field label="Last Name" value={data.info.lname} onInput={(v) => update('info', 'lname', v)} placeholder="e.g. Dela Cruz" />
                <Field label="First Name" value={data.info.fname} onInput={(v) => update('info', 'fname', v)} placeholder="e.g. Juan" />
                <Field label="Middle Name" value={data.info.mname} onInput={(v) => update('info', 'mname', v)} placeholder="e.g. Santos" />
                <Field label="LRN (Learner Reference Number)" value={data.info.lrn} onInput={(v) => update('info', 'lrn', v)} placeholder="12-digit LRN" />
                <Field label="Date of Birth (MM/DD/YYYY)" value={data.info.birthdate} onInput={(v) => update('info', 'birthdate', v)} type="date" />
                <Field label="Sex" value={data.info.sex} onInput={(v) => update('info', 'sex', v)} placeholder="Male / Female" />
                <div className="field-group">
                    <label>Status</label>
                    <div className="checkbox-row" onClick={() => update('info', 'irregular', !data.info.irregular)}>
                        <input type="checkbox" checked={data.info.irregular || false} readOnly />
                        <span>Irregular Student</span>
                    </div>
                </div>
                <Field label="Date of SHS Admission (MM/DD/YYYY)" value={data.info.admissionDate} onInput={(v) => update('info', 'admissionDate', v)} type="date" fullSpan />
            </div>
        </div>
    );

    // ── Render Eligibility tab ──
    const renderEligibility = () => {
        const e = data.eligibility;
        return (
            <div className="editor-section">
                <h3><span className="section-icon"><CheckIcon /></span> Eligibility for SHS Enrolment</h3>

                <div className="form-grid">
                    <div className="field-group">
                        <label>High School Completer</label>
                        <div className="checkbox-row" onClick={() => update('eligibility', 'hsCompleter', !e.hsCompleter)}>
                            <input type="checkbox" checked={e.hsCompleter || false} readOnly />
                            <span>High School Completer*</span>
                        </div>
                    </div>
                    <Field label="HS General Average" value={e.hsGenAve} onInput={(v) => update('eligibility', 'hsGenAve', v)} placeholder="e.g. 88.5" />

                    <div className="field-group">
                        <label>Junior High School Completer</label>
                        <div className="checkbox-row" onClick={() => update('eligibility', 'jhsCompleter', !e.jhsCompleter)}>
                            <input type="checkbox" checked={e.jhsCompleter || false} readOnly />
                            <span>JHS Completer</span>
                        </div>
                    </div>
                    <Field label="JHS General Average" value={e.jhsGenAve} onInput={(v) => update('eligibility', 'jhsGenAve', v)} placeholder="e.g. 90.0" />

                    <Field label="Date of Graduation/Completion (MM/DD/YYYY)" value={e.gradDate} onInput={(v) => update('eligibility', 'gradDate', v)} type="date" />
                    <Field label="Name of School" value={e.schoolName} onInput={(v) => update('eligibility', 'schoolName', v)} placeholder="e.g. Capas National High School" />
                    <Field label="School Address" value={e.schoolAddress} onInput={(v) => update('eligibility', 'schoolAddress', v)} placeholder="e.g. Capas, Tarlac" fullSpan />
                </div>

                <div style={{ marginTop: '20px' }}>
                    <div className="form-grid">
                        <div className="field-group">
                            <label>PEPT Passer</label>
                            <div className="checkbox-row" onClick={() => update('eligibility', 'pept', !e.pept)}>
                                <input type="checkbox" checked={e.pept || false} readOnly />
                                <span>PEPT Passer**</span>
                            </div>
                        </div>
                        <Field label="PEPT Rating" value={e.peptRating} onInput={(v) => update('eligibility', 'peptRating', v)} placeholder="Rating" />

                        <div className="field-group">
                            <label>ALS A&E Passer</label>
                            <div className="checkbox-row" onClick={() => update('eligibility', 'als', !e.als)}>
                                <input type="checkbox" checked={e.als || false} readOnly />
                                <span>ALS A&E Passer***</span>
                            </div>
                        </div>
                        <Field label="ALS Rating" value={e.alsRating} onInput={(v) => update('eligibility', 'alsRating', v)} placeholder="Rating" />

                        <Field label="Date of Examination/Assessment (MM/DD/YYYY)" value={e.examDate} onInput={(v) => update('eligibility', 'examDate', v)} type="date" />
                        <Field label="Name/Address of Community Learning Center" value={e.clcName} onInput={(v) => update('eligibility', 'clcName', v)} placeholder="CLC Name and Address" />

                        <div className="field-group">
                            <label>Others</label>
                            <div className="checkbox-row" onClick={() => update('eligibility', 'others', !e.others)}>
                                <input type="checkbox" checked={e.others || false} readOnly />
                                <span>Others (Pls. Specify)</span>
                            </div>
                        </div>
                        <Field label="Others — Specify" value={e.othersSpec} onInput={(v) => update('eligibility', 'othersSpec', v)} placeholder="Specify" />
                    </div>
                </div>

                <div style={{ marginTop: '16px', fontSize: '11px', color: '#64748b', lineHeight: 1.5 }}>
                    <div>*High School Completers are students who graduated under the old curriculum (before K-12).</div>
                    <div>**PEPT — Philippine Educational Placement Test</div>
                    <div>***ALS A&E — Alternative Learning System Accreditation and Equivalency</div>
                </div>
            </div>
        );
    };

    // ── Render a Semester tab ──
    const renderSemester = (semKey) => {
        const sem = data[semKey];
        if (!sem) return null;
        const semLabel = semKey.replace('semester', 'Semester ');

        return (
            <>
                <div className="editor-section">
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                        <h3 style={{ margin: 0 }}><span className="section-icon"><BookIcon /></span> {semLabel} — School & Section Info</h3>
                        <div style={{ display: 'flex', gap: '8px' }}>
                            <button
                                className="btn-secondary"
                                style={{ fontSize: '11px', padding: '4px 8px' }}
                                onClick={() => clearGrades(semKey)}
                            >
                                Clear Grades
                            </button>
                            {semKey !== 'semester1' && (
                                <button
                                    className="btn-secondary"
                                    style={{ fontSize: '11px', padding: '4px 8px' }}
                                    onClick={() => copyFromPrevious(semKey)}
                                >
                                    Copy Subj from Prev Sem
                                </button>
                            )}
                        </div>
                    </div>
                    <div className="form-grid">
                        <Field label="School" value={sem.school} onInput={(v) => updateSem(semKey, 'school', v)} placeholder="e.g. Capas Senior High School" />
                        <Field label="School ID" value={sem.schoolId} onInput={(v) => updateSem(semKey, 'schoolId', v)} placeholder="e.g. 306252" />
                        <Field label="Grade Level" value={sem.gradeLevel} onInput={(v) => updateSem(semKey, 'gradeLevel', v)} placeholder="e.g. 11 or 12" />
                        <Field label="School Year" value={sem.sy} onInput={(v) => updateSem(semKey, 'sy', v)} placeholder="e.g. 2024-2025" />
                        <Field label="Semester" value={sem.sem} onInput={(v) => updateSem(semKey, 'sem', v)} placeholder="e.g. 1st or 2nd" />
                        <Field label="Track / Strand" value={sem.trackStrand} onInput={(v) => updateSem(semKey, 'trackStrand', v)} placeholder="e.g. TVL - ICT" />
                        <Field label="Section" value={sem.section} onInput={(v) => updateSem(semKey, 'section', v)} placeholder="e.g. Section Turing" />
                    </div>
                </div>

                <div className="editor-section">
                    <h3><span className="section-icon"><BookIcon /></span> {semLabel} — Scholastic Record</h3>

                    {/* Type-filtered datalists — one per subject type */}
                    {['Core', 'Applied', 'Specialized', 'Other', 'All'].map(type => (
                        <datalist key={type} id={`annex-subjects-${type.toLowerCase()}`}>
                            {(data.annex || [])
                                .filter(a => a.active && (type === 'All' || a.type === type) && a.subject && a.subject.trim() !== '')
                                .map((a, idx) => (
                                    <option key={idx} value={a.subject} />
                                ))}
                        </datalist>
                    ))}

                    <div className="subjects-table-wrap" data-sem={semKey} onKeyDown={(e) => handleTableKeyDown(e, semKey)}>
                        <table className="subjects-table">
                            <thead>
                                <tr>
                                    <th style={{ width: '36px' }}>#</th>
                                    <th style={{ width: '120px' }}>Type</th>
                                    <th>Subject</th>
                                    <th className="center" style={{ width: '80px' }}>Q1</th>
                                    <th className="center" style={{ width: '80px' }}>Q2</th>
                                    <th className="center" style={{ width: '90px' }}>Sem Final</th>
                                    <th className="center" style={{ width: '100px' }}>Action</th>
                                </tr>
                            </thead>
                            <tbody>
                                {sem.subjects.map((s, i) => (
                                    <tr key={i}>
                                        <td className="row-num">{i + 1}</td>
                                        <td>
                                            <select
                                                value={s.type}
                                                onChange={(e) => updateSub(semKey, i, 'type', e.target.value)}
                                                data-idx={i}
                                                data-field="type"
                                            >
                                                <option value="">—</option>
                                                <option value="Core">Core</option>
                                                <option value="Applied">Applied</option>
                                                <option value="Specialized">Specialized</option>
                                                <option value="Other">Other</option>
                                            </select>
                                        </td>
                                        <td>
                                            <input
                                                value={s.subject}
                                                onChange={(e) => updateSub(semKey, i, 'subject', e.target.value)}
                                                placeholder="Subject name"
                                                data-idx={i}
                                                data-field="subject"
                                                list={s.type ? `annex-subjects-${s.type.toLowerCase()}` : 'annex-subjects-all'}
                                            />
                                        </td>
                                        <td>
                                            <input
                                                className="grade-cell"
                                                value={s.q1}
                                                onChange={(e) => updateSub(semKey, i, 'q1', e.target.value)}
                                                placeholder="—"
                                                data-idx={i}
                                                data-field="q1"
                                                type="number"
                                                onWheel={(e) => e.target.blur()} // Prevent accidental scroll change
                                            />
                                        </td>
                                        <td>
                                            <input
                                                className="grade-cell"
                                                value={s.q2}
                                                onChange={(e) => updateSub(semKey, i, 'q2', e.target.value)}
                                                placeholder="—"
                                                data-idx={i}
                                                data-field="q2"
                                                type="number"
                                                onWheel={(e) => e.target.blur()}
                                            />
                                        </td>
                                        <td><input className="grade-cell" value={s.final} placeholder="—" readOnly style={{ opacity: 0.7 }} tabindex="-1" /></td>
                                        <td>
                                            <input
                                                className="grade-cell"
                                                value={s.action}
                                                placeholder="—"
                                                readOnly
                                                tabindex="-1"
                                                style={{
                                                    opacity: 0.85,
                                                    color: s.action === 'PASSED' ? '#4ade80' : s.action === 'FAILED' ? '#f87171' : undefined,
                                                    fontWeight: s.action ? 600 : 400
                                                }}
                                            />
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className="subject-actions">
                        <button className="btn-add-row" onClick={() => addSubject(semKey)}>+ Add Subject</button>
                        <button
                            className="btn-add-row"
                            style={{ background: '#6366f1' }}
                            onClick={() => autoPopulateFromStrand(semKey)}
                        >
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px' }}><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"></path></svg>
                            Auto-Populate (Strand)
                        </button>
                        <button className="btn-remove-row" onClick={() => removeSubject(semKey)}>− Remove Last</button>
                        <button
                            className="btn-remove-row"
                            style={{ background: '#94a3b8', marginLeft: 'auto' }}
                            onClick={() => clearAllSubjects(semKey)}
                        >
                            Clear All Subjects
                        </button>
                    </div>

                    <div className="gen-ave-row">
                        <label>General Average for the Semester:</label>
                        <input value={sem.genAve} onChange={(e) => updateSem(semKey, 'genAve', e.target.value)} placeholder="—" />
                    </div>

                    <div className="form-grid full-span" style={{ marginTop: '16px' }}>
                        <Field label="Remarks" value={sem.remarks} onInput={(v) => updateSem(semKey, 'remarks', v)} placeholder="Enter remarks..." fullSpan />
                    </div>

                    <div className="signatory-grid">
                        <Field label="Prepared by (Adviser Name)" value={sem.adviserName} onInput={(v) => updateSem(semKey, 'adviserName', v)} placeholder="Full name of adviser" />
                        <Field label="Certified True and Correct" value={sem.certName} onInput={(v) => updateSem(semKey, 'certName', v)} placeholder="Authorized person name" />
                        <Field label="Date Checked (MM/DD/YYYY)" value={sem.dateChecked} onInput={(v) => updateSem(semKey, 'dateChecked', v)} type="date" />
                    </div>
                </div>

                {/* Remedial Classes */}
                <div className="editor-section">
                    <div className="remedial-section" style={{ borderTop: 'none', paddingTop: 0, marginTop: 0 }}>
                        <h4>⚠ Remedial Classes — {semLabel}</h4>

                        <div className="form-grid">
                            <Field label="Conducted From (MM/DD/YYYY)" value={sem.remedial.from} onInput={(v) => updateRem(semKey, 'from', v)} type="date" />
                            <Field label="To (MM/DD/YYYY)" value={sem.remedial.to} onInput={(v) => updateRem(semKey, 'to', v)} type="date" />
                            <Field label="School" value={sem.remedial.school} onInput={(v) => updateRem(semKey, 'school', v)} placeholder="School name" />
                            <Field label="School ID" value={sem.remedial.schoolId} onInput={(v) => updateRem(semKey, 'schoolId', v)} placeholder="School ID" />
                        </div>

                        <div className="subjects-table-wrap" style={{ marginTop: '16px' }}>
                            <table className="subjects-table">
                                <thead>
                                    <tr>
                                        <th style={{ width: '36px' }}>#</th>
                                        <th style={{ width: '110px' }}>Type</th>
                                        <th>Subject</th>
                                        <th className="center" style={{ width: '90px' }}>Sem Final</th>
                                        <th className="center" style={{ width: '100px' }}>Remedial Mark</th>
                                        <th className="center" style={{ width: '100px' }}>Recomputed</th>
                                        <th className="center" style={{ width: '90px' }}>Action</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {sem.remedial.subjects.map((rs, ri) => (
                                        <tr key={ri}>
                                            <td className="row-num">{ri + 1}</td>
                                            <td>
                                                <select
                                                    value={rs.type}
                                                    onChange={(e) => updateRemSub(semKey, ri, 'type', e.target.value)}
                                                >
                                                    <option value="">—</option>
                                                    <option value="Core">Core</option>
                                                    <option value="Applied">Applied</option>
                                                    <option value="Specialized">Specialized</option>
                                                    <option value="Other">Other</option>
                                                </select>
                                            </td>
                                            <td>
                                                <input
                                                    value={rs.subject}
                                                    onChange={(e) => updateRemSub(semKey, ri, 'subject', e.target.value)}
                                                    placeholder="Subject name"
                                                    list={rs.type ? `annex-subjects-${rs.type.toLowerCase()}` : 'annex-subjects-all'}
                                                />
                                            </td>
                                            <td><input className="grade-cell" value={rs.semGrade} onChange={(e) => updateRemSub(semKey, ri, 'semGrade', e.target.value)} placeholder="—" /></td>
                                            <td><input className="grade-cell" value={rs.remedialMark} onChange={(e) => updateRemSub(semKey, ri, 'remedialMark', e.target.value)} placeholder="—" /></td>
                                            <td><input className="grade-cell" value={rs.recomputedGrade} onChange={(e) => updateRemSub(semKey, ri, 'recomputedGrade', e.target.value)} placeholder="—" /></td>
                                            <td><input className="grade-cell" value={rs.action} onChange={(e) => updateRemSub(semKey, ri, 'action', e.target.value)} placeholder="—" /></td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>

                        <div className="subject-actions">
                            <button className="btn-add-row" onClick={() => addRemedialSubject(semKey)}>+ Add Remedial</button>
                            <button className="btn-remove-row" onClick={() => removeRemedialSubject(semKey)}>− Remove Last</button>
                        </div>

                        <div className="form-grid" style={{ marginTop: '12px' }}>
                            <Field label="Name of Teacher/Adviser" value={sem.remedial.teacherName} onInput={(v) => updateRem(semKey, 'teacherName', v)} placeholder="Full name" />
                        </div>
                    </div>
                </div>
            </>
        );
    };

    // ── Render Certification tab ──
    const renderCertification = () => {
        const c = data.certification;
        return (
            <div className="editor-section">
                <h3><span className="section-icon"><GradIcon /></span> Certification & Graduation</h3>
                <div className="form-grid">
                    <Field label="Track/Strand Accomplished" value={c.trackStrand} onInput={(v) => update('certification', 'trackStrand', v)} placeholder="e.g. TVL - ICT" fullSpan />
                    <Field label="SHS General Average" value={c.genAve} onInput={(v) => update('certification', 'genAve', v)} placeholder="e.g. 89.50" />
                    <Field label="Date of SHS Graduation (MM/DD/YYYY)" value={c.gradDate} onInput={(v) => update('certification', 'gradDate', v)} type="date" />
                    <Field label="Awards / Honors Received" value={c.awards} onInput={(v) => update('certification', 'awards', v)} placeholder="e.g. With Honors" fullSpan />
                    <Field label="Certified by (School Head Name)" value={c.schoolHead} onInput={(v) => update('certification', 'schoolHead', v)} placeholder="e.g. MARIOLITO G. MAGCALAS, PhD" />
                    <Field label="Certification Date" value={c.certDate} onInput={(v) => update('certification', 'certDate', v)} type="date" />
                    <Field label="Date Issued (MM/DD/YYYY)" value={c.dateIssued} onInput={(v) => update('certification', 'dateIssued', v)} type="date" />
                </div>
                <div className="form-grid" style={{ marginTop: '16px' }}>
                    <Field label="Remarks (purpose of issuing this form)" value={c.remarks} onInput={(v) => update('certification', 'remarks', v)} type="textarea" placeholder="Please indicate the purpose for issuing this permanent record..." fullSpan />
                </div>
            </div>
        );
    };

    // ── Excel Import Logic ──
    const fileInputRef = useRef(null);

    const handleExcelImport = async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        try {
            const buffer = await file.arrayBuffer();
            const wb = read(buffer);
            const newData = JSON.parse(JSON.stringify(data)); // Deep copy
            let semCount = 0;

            // Helper: Find value by label in a specific row
            const findVal = (row, label) => {
                if (!row) return '';
                const idx = row.findIndex(cell => cell && cell.toString().toUpperCase().includes(label));
                if (idx === -1) return '';
                // Search next 10 cells for value
                for (let i = idx + 1; i < idx + 10; i++) {
                    if (row[i]) return row[i].toString().trim();
                }
                return '';
            };

            // Process each sheet
            wb.SheetNames.forEach(sheetName => {
                const ws = wb.Sheets[sheetName];
                const grid = utils.sheet_to_json(ws, { header: 1, defval: '' });

                // Scan rows
                for (let r = 0; r < grid.length; r++) {
                    const row = grid[r];
                    const rowStr = row.join(' ').toUpperCase();

                    // ── LEARNER INFO (Front Sheet) ──
                    if (rowStr.includes("LAST NAME")) {
                        newData.info.lname = findVal(row, "LAST NAME") || newData.info.lname;
                        newData.info.fname = findVal(row, "FIRST NAME") || newData.info.fname;
                        newData.info.mname = findVal(row, "MIDDLE NAME") || newData.info.mname;
                    }
                    if (rowStr.includes("LRN:")) newData.info.lrn = findVal(row, "LRN") || newData.info.lrn;
                    if (rowStr.includes("DATE OF BIRTH")) newData.info.birthdate = findVal(row, "DATE OF BIRTH") || newData.info.birthdate;
                    if (rowStr.includes("SEX:")) newData.info.sex = findVal(row, "SEX") || newData.info.sex;
                    if (rowStr.includes("ADMISSION")) newData.info.admissionDate = findVal(row, "ADMISSION") || newData.info.admissionDate;

                    // ── ELIGIBILITY ──
                    if (rowStr.includes("HIGH SCHOOL COMPLETER")) {
                        if (findVal(row, "GEN. AVE")) newData.eligibility.hsGenAve = findVal(row, "GEN. AVE");
                    }
                    if (rowStr.includes("JUNIOR HIGH SCHOOL COMPLETER")) {
                        // Logic to differentiate HS vs JHS gen ave if on same row
                    }

                    // ── SEMESTER BLOCKS ──
                    // Detect start of a Semester block (School Info)
                    // Template has "SCHOOL:" label.
                    if (rowStr.includes("SCHOOL:") && rowStr.includes("SCHOOL ID:") && semCount < 2) {
                        semCount++;
                        const semKey = `semester${semCount}`;

                        // Capture School Info
                        newData[semKey].school = findVal(row, "SCHOOL:") || newData[semKey].school;
                        newData[semKey].schoolId = findVal(row, "SCHOOL ID") || newData[semKey].schoolId;
                        newData[semKey].gradeLevel = findVal(row, "GRADE LEVEL") || newData[semKey].gradeLevel;
                        newData[semKey].sy = findVal(row, "SY") || newData[semKey].sy;
                        newData[semKey].sem = findVal(row, "SEM") || newData[semKey].sem;

                        // Section/Track usually on next few rows
                        for (let k = 1; k <= 3; k++) {
                            if (grid[r + k]) {
                                const nextRow = grid[r + k];
                                const nextRowStr = nextRow.join(' ').toUpperCase();
                                if (nextRowStr.includes("TRACK")) newData[semKey].trackStrand = findVal(nextRow, "TRACK") || newData[semKey].trackStrand;
                                if (nextRowStr.includes("SECTION")) newData[semKey].section = findVal(nextRow, "SECTION") || newData[semKey].section;
                            }
                        }

                        // ── SUBJECTS TABLE ──
                        // Look for header "SUBJECTS" within next 10 rows
                        let tableStartRow = -1;
                        for (let k = 1; k < 10; k++) {
                            if (grid[r + k] && grid[r + k].join(' ').toUpperCase().includes("SUBJECTS")) {
                                tableStartRow = r + k + 1; // Data starts after header
                                break;
                            }
                        }

                        if (tableStartRow !== -1) {
                            const COL_SUBJ = 8;
                            const COL_Q1 = 45;
                            const COL_Q2 = 50;
                            const COL_FINAL = 55;
                            const COL_ACTION = 60;

                            let subjIdx = 0;
                            let limit = 15; // max subjects

                            for (let k = 0; k < limit; k++) {
                                const currRowIdx = tableStartRow + k;
                                const rowData = grid[currRowIdx];
                                if (!rowData) break;

                                const subjName = rowData[COL_SUBJ];
                                const type = rowData[COL_SUBJ - 8]; // Column A (0) often has Type

                                // Check if end of table (e.g. "General Ave")
                                const rowFull = rowData.join(' ').toUpperCase();
                                if (rowFull.includes("GENERAL AVE") || rowFull.includes("REMARKS")) {
                                    // Capture Gen Ave
                                    if (rowFull.includes("GENERAL AVE")) {
                                        newData[semKey].genAve = rowData[COL_FINAL] || rowData[COL_Q2] || findVal(rowData, "AVE");
                                    }
                                    if (rowFull.includes("REMARKS")) {
                                        newData[semKey].remarks = findVal(rowData, "REMARKS") || newData[semKey].remarks;
                                    }
                                    break;
                                }

                                if (subjName && typeof subjName === 'string') {
                                    if (!newData[semKey].subjects[subjIdx]) {
                                        newData[semKey].subjects[subjIdx] = { ...EMPTY_SUBJECT };
                                    }
                                    newData[semKey].subjects[subjIdx].subject = subjName;
                                    newData[semKey].subjects[subjIdx].type = (typeof rowData[0] === 'string' && rowData[0].length > 2) ? rowData[0] : '';
                                    newData[semKey].subjects[subjIdx].q1 = rowData[COL_Q1] || '';
                                    newData[semKey].subjects[subjIdx].q2 = rowData[COL_Q2] || '';
                                    newData[semKey].subjects[subjIdx].final = rowData[COL_FINAL] || '';
                                    newData[semKey].subjects[subjIdx].action = rowData[COL_ACTION] || '';

                                    subjIdx++;
                                }
                            }
                        }
                    }
                }
            });

            onChange(newData);
            alert('Import Successful! Data loaded from Excel.');

        } catch (error) {
            console.error(error);
            alert('Error importing file: ' + error.message);
        }
        e.target.value = '';
    };

    // ── Render Print Preview (direct HTML from data — no xlsx processing) ──
    const renderPreview = () => (
        <div style={{ padding: '20px', background: '#e2e8f0', display: 'flex', flexDirection: 'column', alignItems: 'center', minHeight: '400px' }}>
            {/* Hidden File Input for Excel Import */}
            <input
                type="file"
                ref={fileInputRef}
                style={{ display: 'none' }}
                accept=".xlsx, .xls"
                onChange={handleExcelImport}
            />

            {/* Toolbar */}
            <div className="no-print" style={{ marginBottom: '16px', position: 'sticky', top: '10px', zIndex: 100, display: 'flex', gap: '10px', alignItems: 'center', background: 'rgba(255,255,255,0.95)', padding: '10px 20px', borderRadius: '8px', boxShadow: '0 2px 8px rgba(0,0,0,0.15)' }}>
                <button
                    className="btn-primary"
                    style={{ fontSize: '14px', padding: '8px 16px', display: 'inline-flex', alignItems: 'center', gap: '6px', background: '#10b981' }}
                    onClick={() => fileInputRef.current?.click()}
                >
                    <BookIcon /> Import Excel
                </button>
                <button
                    className="btn-secondary"
                    style={{ fontSize: '14px', padding: '8px 16px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                    onClick={() => setShowRawData(true)}
                >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="16 18 22 12 16 6"></polyline><polyline points="8 6 2 12 8 18"></polyline></svg>
                    View Raw JSON
                </button>
                <button
                    className="btn-primary"
                    style={{ fontSize: '14px', padding: '8px 16px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                    onClick={handlePrint}
                >
                    <PrintIcon /> Print
                </button>
                <button
                    className="btn-secondary"
                    style={{ fontSize: '14px', padding: '8px 16px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                    onClick={() => setShowChecker(true)}
                >
                    <SearchIcon /> Check Template Placeholders
                </button>
            </div>

            {/* Form 137 Preview — renders directly from data, no loading needed */}
            <div style={{ overflow: 'auto', background: 'white', padding: '20px', maxWidth: '900px', width: '100%', boxShadow: '0 2px 8px rgba(0,0,0,0.1)', borderRadius: '4px' }}>
                <Form137Preview data={data} />
            </div>
        </div>
    );

    // ── Tab content router ──

    const renderAnnex = () => {
        const syncAnnexFromSemesters = () => {
            const newData = { ...data };
            const allSubjects = [];

            // Collect all subjects from all 4 semesters
            [1, 2, 3, 4].forEach(n => {
                const semData = newData[`semester${n}`];
                if (semData && semData.subjects) {
                    semData.subjects.forEach(s => {
                        if (s.subject && s.subject.trim() !== '') {
                            allSubjects.push({ ...s });
                        }
                    });
                }
            });

            if (allSubjects.length === 0) {
                alert('No subjects found in any semester to sync from.');
                return;
            }

            const core = allSubjects.filter(s => s.type === 'Core');
            const applied = allSubjects.filter(s => s.type === 'Applied');
            const specialized = allSubjects.filter(s => s.type === 'Specialized');
            const other = allSubjects.filter(s => s.type !== 'Core' && s.type !== 'Applied' && s.type !== 'Specialized');

            const newAnnex = Array(36).fill(null).map((_, i) => {
                let type = 'Other';
                if (i < 15) type = 'Core';
                else if (i < 22) type = 'Applied';
                else if (i < 31) type = 'Specialized';
                return { type, subject: '', active: true };
            });

            core.forEach((s, i) => { if (i < 15) newAnnex[i] = { ...newAnnex[i], subject: s.subject }; });
            applied.forEach((s, i) => { if (i < 7) newAnnex[15 + i] = { ...newAnnex[15 + i], subject: s.subject }; });
            specialized.forEach((s, i) => { if (i < 9) newAnnex[22 + i] = { ...newAnnex[22 + i], subject: s.subject }; });
            other.forEach((s, i) => { if (i < 5) newAnnex[31 + i] = { ...newAnnex[31 + i], subject: s.subject }; });

            newData.annex = newAnnex;
            onChange(newData);
            scheduleAutoSave(newData);
        };

        const annexData = data.annex || [];

        // Grouping
        const core = annexData.filter((_, i) => i < 15);
        const applied = annexData.filter((_, i) => i >= 15 && i < 22);
        const specialized = annexData.filter((_, i) => i >= 22 && i < 31);
        const other = annexData.filter((_, i) => i >= 31);

        const renderAnnexSection = (label, list, offset) => (
            <div className="annex-section" style={{ marginBottom: '24px' }}>
                <div style={{ marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <div style={{ width: '4px', height: '18px', background: '#3b82f6', borderRadius: '4px' }}></div>
                    <h3 style={{ margin: 0, fontSize: '15px', fontWeight: '700', color: '#e2e8f0', letterSpacing: '0.5px' }}>{label}</h3>
                </div>

                <div className="annex-grid">
                    <div className="annex-header-row" style={{ gridTemplateColumns: '40px 1fr' }}>
                        <div style={{ textAlign: 'center' }}>Active</div>
                        <div>Master Subject Name</div>
                    </div>

                    {list.map((subj, localIdx) => {
                        const i = offset + localIdx;
                        const isActive = subj.active !== false;
                        return (
                            <div key={i} className="annex-row" style={{
                                gridTemplateColumns: '40px 1fr',
                                opacity: isActive ? 1 : 0.4,
                                background: isActive ? 'rgba(255, 255, 255, 0.03)' : 'transparent'
                            }}>
                                <div style={{ display: 'flex', justifyContent: 'center' }}>
                                    <input
                                        type="checkbox"
                                        checked={isActive}
                                        onChange={(e) => updateAnnexSub(i, 'active', e.target.checked)}
                                        style={{ width: '16px', height: '16px', cursor: 'pointer' }}
                                    />
                                </div>
                                <input
                                    placeholder="Enter subject name here..."
                                    value={subj.subject || ''}
                                    onChange={(e) => updateAnnexSub(i, 'subject', e.target.value)}
                                    onKeyDown={(e) => handleTableKeyDown(e, 'annex')}
                                    data-idx={i} data-field="subject"
                                    style={{ width: '100%', textDecoration: isActive ? 'none' : 'line-through' }}
                                />
                            </div>
                        );
                    })}
                </div>
            </div>
        );

        return (
            <div className="page-view-group">
                <div style={{
                    display: 'grid',
                    gridTemplateColumns: '1fr 1fr',
                    gap: '12px',
                    marginBottom: '26px',
                }}>
                    <div style={{
                        background: 'rgba(59, 130, 246, 0.05)',
                        border: '1px solid rgba(59, 130, 246, 0.2)',
                        backdropFilter: 'blur(10px)',
                        color: '#93c5fd',
                        padding: '16px 20px',
                        borderRadius: '12px',
                        fontSize: '13px',
                        display: 'flex',
                        alignItems: 'flex-start',
                        gap: '12px',
                    }}>
                        <div style={{ background: 'rgba(59, 130, 246, 0.2)', padding: '8px', borderRadius: '8px', display: 'flex', flexShrink: 0 }}>
                            <span style={{ fontSize: '18px', lineHeight: 1 }}>①</span>
                        </div>
                        <div style={{ flex: 1 }}>
                            <div style={{ fontWeight: '700', color: '#fff', marginBottom: '4px' }}>Define Subjects Here First</div>
                            <div style={{ opacity: 0.8, lineHeight: 1.5, marginBottom: '10px' }}>Type your subjects in this Master List, or click the button to auto-fill with standard K-12 subjects from the template.</div>
                            <button
                                onClick={loadTemplateSubjects}
                                style={{
                                    background: 'linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)',
                                    border: 'none',
                                    color: '#fff',
                                    padding: '7px 14px',
                                    borderRadius: '8px',
                                    fontSize: '12px',
                                    fontWeight: '600',
                                    cursor: 'pointer',
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: '6px',
                                }}
                            >
                                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>
                                Load from Template
                            </button>
                        </div>
                    </div>
                    <div style={{
                        background: 'rgba(16, 185, 129, 0.05)',
                        border: '1px solid rgba(16, 185, 129, 0.2)',
                        backdropFilter: 'blur(10px)',
                        color: '#6ee7b7',
                        padding: '16px 20px',
                        borderRadius: '12px',
                        fontSize: '13px',
                        display: 'flex',
                        alignItems: 'flex-start',
                        gap: '12px',
                    }}>
                        <div style={{ background: 'rgba(16, 185, 129, 0.2)', padding: '8px', borderRadius: '8px', display: 'flex', flexShrink: 0 }}>
                            <span style={{ fontSize: '18px', lineHeight: 1 }}>②</span>
                        </div>
                        <div>
                            <div style={{ fontWeight: '700', color: '#fff', marginBottom: '4px' }}>Then Enter Grades in Semester Tabs</div>
                            <div style={{ opacity: 0.8, lineHeight: 1.5 }}>Go to Grade 11 / Grade 12 tabs. Click the Subject field and pick from your list. Grades entered there will auto-fill the Excel Annex sheet on export.</div>
                        </div>
                    </div>
                </div>
                <div className="annex-container">
                    {renderAnnexSection('CORE SUBJECTS', core, 0)}
                    {renderAnnexSection('APPLIED SUBJECTS', applied, 15)}
                    {renderAnnexSection('SPECIALIZED SUBJECTS', specialized, 22)}
                    {renderAnnexSection('OTHER SUBJECTS', other, 31)}
                </div>
            </div>
        );
    };

    const renderTabContent = () => {
        switch (activeTab) {
            case 'info': return renderInfo();
            case 'eligibility': return renderEligibility();
            case 'front': return (
                <div className="page-view-group">
                    {renderSemester('semester1')}
                    <div style={{ height: '40px' }}></div>
                    {renderSemester('semester2')}
                </div>
            );
            case 'back': return (
                <div className="page-view-group">
                    {renderSemester('semester3')}
                    <div style={{ height: '40px' }}></div>
                    {renderSemester('semester4')}
                </div>
            );
            case 'annex': return renderAnnex();
            case 'certification': return renderCertification();
            default: return null;
        }
    };

    return (
        <div className="student-editor-container">
            <div className="editor-tabs">
                {TABS.map((tab) => (
                    <button
                        key={tab.key}
                        className={`editor-tab${activeTab === tab.key ? ' active' : ''}`}
                        onClick={() => setActiveTab(tab.key)}
                    >
                        {tab.icon} {tab.label}
                    </button>
                ))}

                <div style={{ flex: 1 }}></div>

                <div className="tab-actions" style={{ display: 'flex', gap: '8px', paddingRight: '10px' }}>
                    <button
                        className="btn-primary"
                        style={{ height: '32px', padding: '0 12px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px', background: '#10b981' }}
                        onClick={handlePrint}
                    >
                        <PrintIcon /> {printActionLabel}
                    </button>
                </div>
            </div>

            <div className="editor-content-area">
                {renderTabContent()}
            </div>

            {isPrinting && (
                <div style={{
                    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
                    background: 'rgba(0,0,0,0.7)', zIndex: 9999,
                    display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                    color: 'white', fontFamily: 'sans-serif'
                }}>
                    <div className="spinner" style={{ marginBottom: '16px', border: '4px solid rgba(255,255,255,0.3)', borderTop: '4px solid white', borderRadius: '50%', width: '40px', height: '40px', animation: 'spin 1s linear infinite' }}></div>
                    <div style={{ fontSize: '24px', marginBottom: '8px' }}>Processing...</div>
                    <div style={{ fontSize: '16px', opacity: 0.9 }}>{isDesktopMode ? 'Opening Excel for Printing...' : 'Opening Excel and sending to printer...'}</div>
                </div>
            )}
            {showChecker && (
                <PlaceholderChecker onClose={() => setShowChecker(false)} />
            )}
            {showRawData && (
                <div className="modal-overlay" style={{ display: 'flex' }}>
                    <div className="modal-content premium-modal" style={{ maxWidth: '800px', width: '90%' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
                            <h2 style={{ margin: 0 }}>Student Raw Data (JSON)</h2>
                            <button className="btn-secondary" onClick={() => setShowRawData(false)}>Close</button>
                        </div>
                        <div style={{ background: '#0f172a', padding: '15px', borderRadius: '8px', overflow: 'auto', maxHeight: '60vh' }}>
                            <pre style={{ color: '#4ade80', fontSize: '12px', margin: 0 }}>
                                {JSON.stringify(data, null, 2)}
                            </pre>
                        </div>
                        <div style={{ marginTop: '15px', display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
                            <button
                                className="btn-primary"
                                onClick={() => {
                                    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
                                    alert('Copied to clipboard!');
                                }}
                            >
                                Copy to Clipboard
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}

export default StudentEditor;
