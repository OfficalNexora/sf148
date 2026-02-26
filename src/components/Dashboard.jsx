import React, { useState, useEffect, useCallback } from 'react';
import Sidebar from './Sidebar';
import StudentEditor from './StudentEditor';
import UserManagement from './UserManagement';
import Modals from './Modals';
import SyncInbox from './SyncInbox';
import db from '../services/db';
import syncService from '../services/syncService';

function Dashboard({ userRole, onLogout }) {
    const [structure, setStructure] = useState(null);
    const [isEditMode, setIsEditMode] = useState(false);
    const [currentStudent, setCurrentStudent] = useState(null);
    const [currentStudentData, setCurrentStudentData] = useState(null);
    const [status, setStatus] = useState('Ready');
    const [statusColor, setStatusColor] = useState('inherit');
    const [showUserManagement, setShowUserManagement] = useState(false);
    const [searchQuery, setSearchQuery] = useState('');

    // Modal states
    const [alertModal, setAlertModal] = useState({ show: false, message: '' });
    const [promptModal, setPromptModal] = useState({ show: false, title: '', resolve: null });
    const [confirmModal, setConfirmModal] = useState({ show: false, message: '', resolve: null });
    const [shareModal, setShareModal] = useState(false);

    // Online Sync state (admin only)
    const [pendingSyncs, setPendingSyncs] = useState([]);
    const [showSyncInbox, setShowSyncInbox] = useState(false);
    const [sendToAdminName, setSendToAdminName] = useState('');
    const [isSending, setIsSending] = useState(false);

    // Load structure on mount
    useEffect(() => {
        loadStructure();
    }, []);

    // Mount Firestore listener for admin
    useEffect(() => {
        if (userRole !== 'admin') return;
        const unsubscribe = syncService.listenForPendingSyncs((requests) => {
            setPendingSyncs(requests);
        });
        return () => unsubscribe();
    }, [userRole]);

    const loadStructure = async () => {
        const data = await db.getStructure();
        setStructure(data);
    };

    const saveStructure = async (newStructure) => {
        await db.updateStructure(newStructure);
        setStructure(newStructure);
    };

    // Custom dialog helpers
    const showAlert = useCallback((message) => {
        return new Promise((resolve) => {
            setAlertModal({ show: true, message, resolve });
        });
    }, []);

    const showPrompt = useCallback((title) => {
        return new Promise((resolve) => {
            setPromptModal({ show: true, title, resolve });
        });
    }, []);

    const showConfirm = useCallback((message) => {
        return new Promise((resolve) => {
            setConfirmModal({ show: true, message, resolve });
        });
    }, []);

    // Student loading
    const loadStudent = async (id, name) => {
        setShowUserManagement(false);
        setCurrentStudent({ id, name });

        let data = await db.loadStudent(id);
        if (!data) {
            data = generateEmptyRecord(id, name);
        }
        setCurrentStudentData(data);
        setStatus('Ready');
    };

    // Generate empty student record
    const generateEmptyRecord = (id, name) => {
        let lname = '', fname = '';
        if (name && name.includes(',')) {
            [lname, fname] = name.split(',').map(s => s.trim());
        } else if (name) {
            fname = name.trim();
        }

        const makeSem = () => ({
            school: '', schoolId: '', gradeLevel: '', sy: '', sem: '',
            trackStrand: '', section: '',
            subjects: Array.from({ length: 9 }, () => ({ type: '', subject: '', q1: '', q2: '', final: '', action: '' })),
            genAve: '', remarks: '', adviserName: '', certName: '', dateChecked: '',
            remedial: {
                from: '', to: '', school: '', schoolId: '',
                subjects: Array.from({ length: 4 }, () => ({ type: '', subject: '', semGrade: '', remedialMark: '', recomputedGrade: '', action: '' })),
                teacherName: ''
            }
        });

        return {
            info: { lname, fname, mname: '', lrn: '', sex: '', birthdate: '', admissionDate: '', irregular: false },
            eligibility: {
                hsCompleter: false, hsGenAve: '', jhsCompleter: false, jhsGenAve: '',
                gradDate: '', schoolName: '', schoolAddress: '',
                pept: false, peptRating: '', als: false, alsRating: '',
                examDate: '', clcName: '', others: false, othersSpec: ''
            },
            semester1: makeSem(),
            semester2: makeSem(),
            semester3: makeSem(),
            semester4: makeSem(),
            annex: [
                ...Array.from({ length: 15 }, () => ({ type: 'Core', subject: '', active: true })),
                ...Array.from({ length: 7 }, () => ({ type: 'Applied', subject: '', active: true })),
                ...Array.from({ length: 9 }, () => ({ type: 'Specialized', subject: '', active: true })),
                ...Array.from({ length: 5 }, () => ({ type: 'Other', subject: '', active: true }))
            ],
            certification: { trackStrand: '', genAve: '', awards: '', gradDate: '', schoolHead: '', certDate: '', remarks: '', dateIssued: '' }
        };
    };

    // Save student data
    const saveStudentData = useCallback(async (data) => {
        if (!currentStudent) return;

        setStatus('Saving...');
        setStatusColor('orange');

        await db.saveStudent(currentStudent.id, data);
        setCurrentStudentData(data);

        // Sync metadata (e.g. irregular status) to structure for Sidebar filtering
        setStructure(prevStructure => {
            let changed = false;
            const newStructure = { ...prevStructure };

            const updateRecursive = (obj) => {
                if (!obj) return;
                Object.keys(obj).forEach(key => {
                    if (Array.isArray(obj[key])) {
                        const idx = obj[key].findIndex(s => s.id === currentStudent.id);
                        if (idx !== -1) {
                            const student = obj[key][idx];
                            const newIrregular = !!data.info.irregular;
                            const newLrn = data.info.lrn;
                            if (student.irregular !== newIrregular || student.lrn !== newLrn) {
                                obj[key] = [...obj[key]]; // Shallow copy the array
                                obj[key][idx] = { ...student, irregular: newIrregular, lrn: newLrn };
                                changed = true;
                            }
                        }
                    } else if (typeof obj[key] === 'object' && obj[key] !== null) {
                        updateRecursive(obj[key]);
                    }
                });
            };

            updateRecursive(newStructure);

            if (changed) {
                db.updateStructure(newStructure);
                return newStructure;
            }
            return prevStructure;
        });

        setStatus('Saved');
        setStatusColor('green');

        setTimeout(() => {
            setStatus('Ready');
            setStatusColor('inherit');
        }, 2000);
    }, [currentStudent]);

    // --- Toggle irregular status directly from sidebar ---
    const toggleIrregular = async (studentId) => {
        let found = false;
        const newStructure = { ...structure };

        const findAndToggle = (obj) => {
            if (found || !obj) return;
            Object.keys(obj).forEach(key => {
                if (Array.isArray(obj[key])) {
                    const idx = obj[key].findIndex(s => s.id === studentId);
                    if (idx !== -1) {
                        obj[key] = [...obj[key]];
                        obj[key][idx] = { ...obj[key][idx], irregular: !obj[key][idx].irregular };
                        found = true;
                    }
                } else if (typeof obj[key] === 'object' && obj[key] !== null) {
                    findAndToggle(obj[key]);
                }
            });
        };
        findAndToggle(newStructure);

        if (found) {
            await saveStructure(newStructure);

            // Also update the student record data if it's currently loaded
            if (currentStudent && currentStudent.id === studentId && currentStudentData) {
                const updatedData = {
                    ...currentStudentData,
                    info: { ...currentStudentData.info, irregular: !currentStudentData.info.irregular }
                };
                setCurrentStudentData(updatedData);
                await db.saveStudent(studentId, updatedData);
            }
        }
    };

    // Structure management
    const addGrade = async () => {
        const name = await showPrompt('Enter Grade Level Name (e.g., Grade 11):');
        if (!name) return;
        if (structure[name]) {
            await showAlert('Grade Level already exists.');
            return;
        }
        const newStructure = { ...structure, [name]: {} };
        await saveStructure(newStructure);
    };

    const addStrand = async (path) => {
        const name = await showPrompt('Enter Strand Name (e.g., TVL - ICT):');
        if (!name) return;
        if (structure[path[0]][name]) {
            await showAlert('Strand already exists.');
            return;
        }
        const newStructure = { ...structure };
        newStructure[path[0]] = { ...newStructure[path[0]], [name]: {} };
        await saveStructure(newStructure);
    };

    const addSection = async (path) => {
        const name = await showPrompt('Enter Section Name (e.g., Section A):');
        if (!name) return;
        if (structure[path[0]][path[1]][name]) {
            await showAlert('Section already exists.');
            return;
        }
        const newStructure = { ...structure };
        newStructure[path[0]][path[1]] = { ...newStructure[path[0]][path[1]], [name]: [] };
        await saveStructure(newStructure);
    };

    const addStudent = async (path) => {
        const name = await showPrompt('Enter Student Name (Last Name, First Name):');
        if (!name) return;
        const id = Date.now().toString();
        const newStructure = { ...structure };
        newStructure[path[0]][path[1]][path[2]] = [
            ...newStructure[path[0]][path[1]][path[2]],
            { id, name, irregular: false }
        ];
        await saveStructure(newStructure);
    };

    const deleteItem = async (path) => {
        if (!await showConfirm('Are you sure you want to delete this item? This cannot be undone.')) return;

        const newStructure = JSON.parse(JSON.stringify(structure));

        if (path.length === 1) {
            delete newStructure[path[0]];
        } else if (path.length === 2) {
            delete newStructure[path[0]][path[1]];
        } else if (path.length === 3) {
            delete newStructure[path[0]][path[1]][path[2]];
        } else if (path.length === 4) {
            newStructure[path[0]][path[1]][path[2]].splice(path[3], 1);
        }

        await saveStructure(newStructure);
    };

    // --- Export functions (browser-compatible) ---
    const exportToExcel = async () => {
        if (!currentStudentData) return;

        if (db.isElectron()) {
            // Electron: use IPC for full template-based export
            const ipcRenderer = window.require('electron').ipcRenderer;
            const result = await ipcRenderer.invoke('export-excel', currentStudentData);
            if (result.success) {
                setShareModal(false);
                await showAlert('Exported to Excel successfully!');
            } else if (!result.cancelled) {
                await showAlert('Export failed: ' + (result.error || 'Unknown error'));
            }
        } else {
            // Browser: generate Excel using SheetJS
            try {
                const XLSX = await import('xlsx');
                const wb = XLSX.utils.book_new();
                const info = currentStudentData.info;

                // Build a summary sheet with student info
                const infoRows = [
                    ['FORM 137 - Student Permanent Record'],
                    [],
                    ['Last Name', info.lname],
                    ['First Name', info.fname],
                    ['Middle Name', info.mname],
                    ['LRN', info.lrn],
                    ['Sex', info.sex],
                    ['Date of Birth', info.birthdate],
                    ['Irregular', info.irregular ? 'Yes' : 'No'],
                    []
                ];

                // Add semester data
                ['semester1', 'semester2', 'semester3', 'semester4'].forEach((semKey, idx) => {
                    const sem = currentStudentData[semKey];
                    if (!sem) return;
                    infoRows.push([`--- Semester ${idx + 1} ---`]);
                    infoRows.push(['School', sem.school]);
                    infoRows.push(['School Year', sem.sy]);
                    infoRows.push(['Grade Level', sem.gradeLevel]);
                    infoRows.push(['Track/Strand', sem.trackStrand]);
                    infoRows.push(['Section', sem.section]);
                    infoRows.push([]);
                    infoRows.push(['Type', 'Subject', 'Q1', 'Q2', 'Final', 'Action']);

                    (sem.subjects || []).forEach(subj => {
                        if (subj.subject) {
                            infoRows.push([subj.type, subj.subject, subj.q1, subj.q2, subj.final, subj.action]);
                        }
                    });
                    infoRows.push(['General Average', sem.genAve]);
                    infoRows.push(['Remarks', sem.remarks]);
                    infoRows.push([]);
                });

                const ws = XLSX.utils.aoa_to_sheet(infoRows);
                XLSX.utils.book_append_sheet(wb, ws, 'Form 137');

                const filename = `Form137_${info.lname || 'Student'}.xlsx`;
                XLSX.writeFile(wb, filename);
                setShareModal(false);
                await showAlert('Exported to Excel successfully!');
            } catch (err) {
                console.error('Excel export error:', err);
                await showAlert('Export failed: ' + err.message);
            }
        }
    };

    const exportToFile = async () => {
        if (!currentStudentData) {
            await showAlert('Please select a student first.');
            return;
        }

        if (db.isElectron()) {
            const ipcRenderer = window.require('electron').ipcRenderer;
            const filePath = await ipcRenderer.invoke('save-file-dialog', `Form137_${currentStudentData.info.lname}.json`);
            if (!filePath) return;
            const path = window.require('path');
            const result = await ipcRenderer.invoke('export-file', {
                path: path.dirname(filePath),
                filename: path.basename(filePath),
                content: JSON.stringify(currentStudentData, null, 2)
            });
            await showAlert(result.success ? 'Export successful!' : 'Export failed: ' + result.error);
        } else {
            const filename = `Form137_${currentStudentData.info.lname || 'Student'}.json`;
            await db.exportStudentToFile(currentStudentData, filename);
            await showAlert('Export successful! File downloaded.');
        }
        setShareModal(false);
    };

    const importStudent = async () => {
        let importedData = null;

        if (db.isElectron()) {
            const ipcRenderer = window.require('electron').ipcRenderer;
            const filePath = await ipcRenderer.invoke('open-file-dialog');
            if (!filePath) return;

            if (filePath.endsWith('.xlsx')) {
                importedData = await ipcRenderer.invoke('import-excel-form', filePath);
            } else {
                const content = await ipcRenderer.invoke('read-file', filePath);
                if (content) {
                    try { importedData = JSON.parse(content); } catch (e) { }
                }
            }
        } else {
            // Browser mode: use file picker
            importedData = await db.importStudentFromFile();
        }

        if (!importedData) {
            await showAlert('Failed to import file. Ensure it is a valid Form 137 JSON.');
            return;
        }

        // --- SYNC LOGIC ---
        const importedLrn = importedData.info.lrn;
        let existingStudent = null;

        const findExistingByLrn = (obj) => {
            if (existingStudent || !obj) return;
            Object.keys(obj).forEach(key => {
                if (Array.isArray(obj[key])) {
                    const match = obj[key].find(s => s.lrn === importedLrn);
                    if (match) existingStudent = match;
                } else {
                    findExistingByLrn(obj[key]);
                }
            });
        };

        if (importedLrn && structure) {
            findExistingByLrn(structure);
        }

        if (existingStudent) {
            const confirmed = await showConfirm(`Student with LRN ${importedLrn} already exists (${existingStudent.name}). Do you want to SYNC (update) this record?`);
            if (confirmed) {
                const oldData = await db.loadStudent(existingStudent.id);
                const mergedData = {
                    ...oldData,
                    ...importedData,
                    info: { ...oldData.info, ...importedData.info },
                    eligibility: { ...oldData.eligibility, ...importedData.eligibility }
                };
                await saveStudentData(mergedData);
                await loadStudent(existingStudent.id, existingStudent.name);
                await showAlert('Sync complete! Record updated.');
                return;
            }
        }

        // Otherwise, import as NEW student
        const name = `${importedData.info.lname}, ${importedData.info.fname}`;
        const confirmedNew = await showConfirm(`Import ${name} as a new record?`);
        if (!confirmedNew) return;

        const id = Date.now().toString();
        setCurrentStudentData(importedData);
        setCurrentStudent({ id, name });
        setShowUserManagement(false);
        await showAlert('Import successful! (Note: New students are not yet linked to the sidebar structure. Please save to finalize.)');
    };

    const handlePrint = () => {
        if (!currentStudentData) {
            showAlert('Please select a student first.');
            return;
        }
        window.print();
    };

    // --- Database Sync ---
    const exportSyncDb = async () => {
        const result = await db.exportSyncFile();
        if (result && result.success) {
            await showAlert('Sync file exported! Share it with the Admin to merge data.');
        } else if (result && !result.success) {
            await showAlert('Export cancelled or failed.');
        }
    };

    const importSyncDb = async () => {
        const result = await db.importSyncFile();
        if (result && result.success) {
            await showAlert(`Sync complete! Merged ${result.count} student record(s).`);
            loadStructure();
        } else if (result && result.error) {
            await showAlert('Import failed: ' + result.error);
        }
    };

    return (
        <div id="dashboard-view">
            <Sidebar
                structure={structure}
                isEditMode={isEditMode}
                onToggleEditMode={() => setIsEditMode(!isEditMode)}
                onRefresh={loadStructure}
                onStudentSelect={loadStudent}
                onAddGrade={addGrade}
                onAddStrand={addStrand}
                onAddSection={addSection}
                onAddStudent={addStudent}
                onDeleteItem={deleteItem}
                onToggleIrregular={toggleIrregular}
                searchQuery={searchQuery}
                onSearchChange={setSearchQuery}
                userRole={userRole}
                onUserManagement={() => {
                    setShowUserManagement(true);
                    setCurrentStudent(null);
                    setCurrentStudentData(null);
                }}
                onShare={() => setShareModal(true)}
                onLogout={onLogout}
                pendingSyncCount={pendingSyncs.length}
                onOpenSyncInbox={() => setShowSyncInbox(true)}
            />

            <main className="workspace">
                <div className="toolbar">
                    <div className="toolbar-left">

                        <button
                            className="btn-primary"
                            style={{ width: 'auto', padding: '5px 15px', background: '#28a745' }}
                            onClick={importStudent}
                        >
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="12" y1="18" x2="12" y2="12"></line><polyline points="9 15 12 12 15 15"></polyline></svg>
                            Import Student
                        </button>
                    </div>
                    <div className="status-bar" style={{ color: statusColor }}>{status}</div>
                </div>

                <div className="editor-viewport">
                    <div id="editor-content">
                        {showUserManagement ? (
                            <UserManagement showAlert={showAlert} />
                        ) : currentStudentData ? (
                            <StudentEditor
                                data={currentStudentData}
                                onChange={(newData) => {
                                    setCurrentStudentData(newData);
                                    setStatus('Unsaved...');
                                    setStatusColor('orange');
                                }}
                                onSave={saveStudentData}
                            />
                        ) : (
                            <div id="welcome-screen" style={{ textAlign: 'center', color: '#888', marginTop: '50px' }}>
                                <h3>Welcome to Form 137 Maker</h3>
                                <p>Select a student from the sidebar to view or edit their record.</p>
                            </div>
                        )}
                    </div>
                </div>
            </main>

            {/* Share & Sync Modal */}
            {shareModal && (
                <div className="modal-overlay" style={{ display: 'flex' }}>
                    <div className="modal-content premium-modal center-text">
                        <h2 style={{ marginTop: 0, marginBottom: '5px' }}>Share & Sync</h2>
                        <p style={{ color: '#cbd5e1', marginBottom: '20px', fontSize: '13px' }}>Export this student's record, or sync the entire database.</p>

                        <p style={{ color: '#94a3b8', fontSize: '11px', textTransform: 'uppercase', letterSpacing: '1px', marginBottom: '10px' }}>Student Record Export</p>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginBottom: '20px' }}>
                            <button
                                className="btn-primary"
                                style={{ background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)', padding: '12px' }}
                                onClick={exportToExcel}
                            >
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                                Export to Excel (.xlsx)
                            </button>
                            <button
                                className="btn-primary"
                                style={{ background: 'linear-gradient(135deg, #3b82f6 0%, #2563eb 100%)', padding: '12px' }}
                                onClick={exportToFile}
                            >
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                                Export to JSON (.json)
                            </button>
                        </div>

                        <p style={{ color: '#94a3b8', fontSize: '11px', textTransform: 'uppercase', letterSpacing: '1px', marginBottom: '10px' }}>Online Database Sync</p>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                            {/* Send to Admin (for teachers) */}
                            <div style={{ display: 'flex', gap: '8px' }}>
                                <input
                                    type="text"
                                    placeholder="Your name (e.g. Ms. Santos)"
                                    value={sendToAdminName}
                                    onChange={e => setSendToAdminName(e.target.value)}
                                    style={{
                                        flex: 1,
                                        padding: '8px 10px',
                                        borderRadius: '6px',
                                        border: '1px solid rgba(255,255,255,0.15)',
                                        background: 'rgba(255,255,255,0.07)',
                                        color: 'white',
                                        fontSize: '12px'
                                    }}
                                />
                                <button
                                    className="btn-primary"
                                    disabled={isSending || !sendToAdminName.trim()}
                                    style={{
                                        padding: '8px 14px',
                                        background: isSending ? '#555' : 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
                                        whiteSpace: 'nowrap',
                                        fontSize: '12px'
                                    }}
                                    onClick={async () => {
                                        setIsSending(true);
                                        try {
                                            // Gather local DB snapshot
                                            const structureData = await db.getStructure();
                                            const bdb = db._getBrowserDB ? db._getBrowserDB() : null;
                                            let records = [];
                                            if (bdb) {
                                                await bdb.ready;
                                                const rawRecords = await bdb._getAll('records');
                                                records = rawRecords.map(r => r.data);
                                            }
                                            const syncData = { structure: structureData, records };
                                            const result = await syncService.submitSync(sendToAdminName.trim(), syncData);
                                            if (result.success) {
                                                setShareModal(false);
                                                setSendToAdminName('');
                                                await showAlert('Sent! The admin will receive your data instantly.');
                                            } else {
                                                await showAlert('Failed to send: ' + result.error);
                                            }
                                        } finally {
                                            setIsSending(false);
                                        }
                                    }}
                                >
                                    {isSending ? 'Sending...' : 'Send to Admin'}
                                </button>
                            </div>
                            {/* Offline fallback */}
                            <button
                                className="btn-primary"
                                style={{ background: 'rgba(255,255,255,0.08)', padding: '10px', fontSize: '12px', color: '#94a3b8' }}
                                onClick={exportSyncDb}
                            >
                                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px' }}><polyline points="8 17 12 21 16 17"></polyline><line x1="12" y1="12" x2="12" y2="21"></line><path d="M20.88 18.09A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.29"></path></svg>
                                Download Sync File (offline fallback)
                            </button>
                            {userRole === 'admin' && (
                                <button
                                    className="btn-primary"
                                    style={{ background: 'linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%)', padding: '10px', fontSize: '12px' }}
                                    onClick={() => { setShareModal(false); setShowSyncInbox(true); }}
                                >
                                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px' }}><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg>
                                    Open Sync Inbox {pendingSyncs.length > 0 && `(${pendingSyncs.length} pending)`}
                                </button>
                            )}
                        </div>

                        <button
                            className="btn-secondary full-width"
                            style={{ marginTop: '20px' }}
                            onClick={() => setShareModal(false)}
                        >
                            Cancel
                        </button>
                    </div>
                </div>
            )}

            {/* Sync Inbox Modal (Admin Only) */}
            {showSyncInbox && (
                <div className="modal-overlay" style={{ display: 'flex' }}>
                    <div className="modal-content premium-modal" style={{ maxWidth: '420px', width: '100%' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                            <h2 style={{ margin: 0, fontSize: '17px' }}>Sync Inbox</h2>
                            <button
                                onClick={() => setShowSyncInbox(false)}
                                style={{ background: 'none', border: 'none', color: '#94a3b8', cursor: 'pointer', fontSize: '18px' }}
                            >✕</button>
                        </div>
                        <SyncInbox
                            requests={pendingSyncs}
                            onMerge={() => { loadStructure(); }}
                            onDismiss={() => { }}
                        />
                    </div>
                </div>
            )}

            <Modals
                alertModal={alertModal}
                onAlertClose={() => {
                    alertModal.resolve?.();
                    setAlertModal({ show: false, message: '' });
                }}
                promptModal={promptModal}
                onPromptSubmit={(value) => {
                    promptModal.resolve?.(value);
                    setPromptModal({ show: false, title: '', resolve: null });
                }}
                onPromptCancel={() => {
                    promptModal.resolve?.(null);
                    setPromptModal({ show: false, title: '', resolve: null });
                }}
                confirmModal={confirmModal}
                onConfirmYes={() => {
                    confirmModal.resolve?.(true);
                    setConfirmModal({ show: false, message: '', resolve: null });
                }}
                onConfirmNo={() => {
                    confirmModal.resolve?.(false);
                    setConfirmModal({ show: false, message: '', resolve: null });
                }}
            />
        </div>
    );
}

export default Dashboard;
