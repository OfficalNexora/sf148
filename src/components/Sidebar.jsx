import React, { useMemo } from 'react';

function Sidebar({
    structure,
    isEditMode,
    onToggleEditMode,
    onRefresh,
    onStudentSelect,
    onAddGrade,
    onAddStrand,
    onAddSection,
    onAddStudent,
    onDeleteItem,
    onToggleIrregular,
    searchQuery,
    onSearchChange,
    userRole,
    onUserManagement,
    onShare,
    onLogout,
    pendingSyncCount = 0,
    onOpenSyncInbox
}) {
    // Determine if searching specifically for irregular students
    const isIrregularSearch = useMemo(() => {
        if (!searchQuery) return false;
        const q = searchQuery.toLowerCase().trim();
        return q === 'irr' || q === 'irregular' || q.startsWith('irr');
    }, [searchQuery]);

    // Pre-calculate paths that need to be expanded when searching
    // This runs ONCE per keystroke, preventing O(N^2) lag during rendering
    const expandedPaths = useMemo(() => {
        const expanded = new Set();
        if (!searchQuery || !structure) return expanded;

        const query = searchQuery.toLowerCase().trim();

        const traverse = (obj, currentPath) => {
            let hasMatchInSubtree = false;
            if (!obj) return false;

            Object.keys(obj).forEach(key => {
                const path = [...currentPath, key];
                const pathKey = path.join('|');

                if (Array.isArray(obj[key])) {
                    // Check students in this section
                    let sectionMatch = false;
                    for (const s of obj[key]) {
                        if (isIrregularSearch) {
                            if (s.irregular) { sectionMatch = true; break; }
                        } else {
                            if (s.name.toLowerCase().includes(query)) { sectionMatch = true; break; }
                        }
                    }
                    if (sectionMatch) {
                        hasMatchInSubtree = true;
                        expanded.add(pathKey);
                    }
                } else {
                    // Recursively check deeper levels
                    const childMatch = traverse(obj[key], path);
                    if (childMatch) {
                        hasMatchInSubtree = true;
                        expanded.add(pathKey);
                    }
                }
            });
            return hasMatchInSubtree;
        };

        traverse(structure, []);
        return expanded;
    }, [structure, searchQuery, isIrregularSearch]);

    const highlightText = (text, query) => {
        if (!query || !text) return text;
        const parts = text.split(new RegExp(`(${query})`, 'gi'));
        return parts.map((part, i) =>
            part.toLowerCase() === query.toLowerCase()
                ? <span key={i} className="search-highlight-text">{part}</span>
                : part
        );
    };

    const renderTree = (obj, path = []) => {
        if (!obj) return null;

        const keys = Object.keys(obj);
        const depth = path.length;

        // Sorting: Grades usually 11, 12. 
        // We can sort keys if needed, but let's stick to object order for now.

        return keys.map((key) => {
            const currentPath = [...path, key];
            const pathKey = currentPath.join('|');
            const value = obj[key];
            const isExpanded = searchQuery ? expandedPaths.has(pathKey) : false;

            // Determine if this is a leaf node (Section with students array)
            const isSection = Array.isArray(value);

            return (
                <li key={pathKey} className="tree-node">
                    <details open={isExpanded}>
                        <summary className="tree-summary">
                            <span className="summary-content">
                                <span className="label">
                                    {depth === 0 ? `Grade ${key}` : key}
                                </span>
                                {isEditMode && (
                                    <span className="action-buttons">
                                        <button
                                            onClick={(e) => {
                                                e.preventDefault();
                                                e.stopPropagation();
                                                if (depth === 0) onAddStrand(currentPath);
                                                else if (depth === 1) onAddSection(currentPath);
                                                else if (isSection) onAddStudent(currentPath);
                                            }}
                                            className="action-btn add"
                                            title="Add Sub-item"
                                        >
                                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M12 5v14M5 12h14" /></svg>
                                        </button>
                                        <button
                                            onClick={(e) => { e.preventDefault(); e.stopPropagation(); onDeleteItem(currentPath); }}
                                            className="action-btn delete"
                                            title="Delete Item"
                                        >
                                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M18 6L6 18M6 6l12 12" /></svg>
                                        </button>
                                    </span>
                                )}
                            </span>
                        </summary>
                        <ul className="tree-list">
                            {isSection ? (
                                // Render students
                                value.map((student, index) => {
                                    let isMatch = false;
                                    if (isIrregularSearch) {
                                        isMatch = !!student.irregular;
                                    } else if (searchQuery) {
                                        isMatch = student.name.toLowerCase().includes(searchQuery.toLowerCase());
                                    }

                                    if (searchQuery && !isMatch) return null;

                                    return (
                                        <li
                                            key={student.id}
                                            className={`student-node ${isMatch ? 'search-highlight' : ''}`}
                                            onClick={() => onStudentSelect(student.id, student.name)}
                                        >
                                            <div className="student-content">
                                                <span className="student-name">
                                                    {highlightText(student.name, searchQuery)}
                                                    {student.irregular && (
                                                        <span className="irregular-badge">IRR</span>
                                                    )}
                                                </span>
                                                {isEditMode && (
                                                    <span className="student-actions">
                                                        <button
                                                            onClick={(e) => {
                                                                e.stopPropagation();
                                                                onToggleIrregular(student.id);
                                                            }}
                                                            className={`status-btn ${student.irregular ? 'active' : ''}`}
                                                            title={student.irregular ? 'Remove Irregular status' : 'Mark as Irregular'}
                                                        >
                                                            <svg width="12" height="12" viewBox="0 0 24 24" fill={student.irregular ? 'currentColor' : 'none'} stroke="currentColor" strokeWidth="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" /></svg>
                                                        </button>
                                                        <button
                                                            onClick={(e) => {
                                                                e.stopPropagation();
                                                                onDeleteItem([...currentPath, index]);
                                                            }}
                                                            className="delete-btn"
                                                        >
                                                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M18 6L6 18M6 6l12 12" /></svg>
                                                        </button>
                                                    </span>
                                                )}
                                            </div>
                                        </li>
                                    );
                                })
                            ) : (
                                renderTree(value, currentPath)
                            )}
                        </ul>
                    </details>
                </li>
            );
        });
    };

    return (
        <aside className="sidebar">
            <div
                className="brand-content"
                style={{ padding: '15px 10px', borderBottom: '1px solid rgba(255,255,255,0.1)', textAlign: 'center' }}
            >
                <img
                    src="/assets/images/capas_senior_high_school.jpg"
                    alt="School Logo"
                    className="brand-logo"
                    style={{ width: '80px', height: '80px', borderWidth: '2px', borderRadius: '50%', marginBottom: '10px' }}
                />
                <h1 style={{ fontSize: '13px', marginTop: '0', lineHeight: 1.2, color: 'white', fontWeight: 'bold', textTransform: 'uppercase', letterSpacing: '1px' }}>
                    Capas Senior High<br /><span style={{ opacity: 0.8, fontSize: '11px' }}>Form 137 Records</span>
                </h1>
            </div>

            <div style={{ padding: '10px', borderBottom: '1px solid rgba(255,255,255,0.1)', display: 'flex', gap: '5px' }}>
                <button
                    className="btn-primary"
                    onClick={onToggleEditMode}
                    style={{
                        padding: '5px',
                        fontSize: '12px',
                        margin: 0,
                        background: isEditMode ? '#28a745' : undefined,
                        display: 'flex', alignItems: 'center', gap: '5px', justifyContent: 'center'
                    }}
                >
                    {isEditMode ? (
                        <>
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M20 6L9 17l-5-5" /></svg>
                            Done Editing
                        </>
                    ) : (
                        <>
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                            Edit Structure
                        </>
                    )}
                </button>
                <button
                    className="btn-primary"
                    onClick={onRefresh}
                    style={{ padding: '5px', fontSize: '12px', margin: 0, width: 'auto' }}
                    title="Refresh Structure"
                >
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
                </button>
            </div>

            <div style={{ padding: '10px', borderBottom: '1px solid rgba(255,255,255,0.1)' }}>
                <input
                    type="text"
                    placeholder='Search student, section, or "irregular"...'
                    value={searchQuery}
                    onChange={(e) => onSearchChange(e.target.value)}
                    style={{
                        width: '100%',
                        padding: '8px',
                        borderRadius: '4px',
                        border: 'none',
                        background: 'rgba(255,255,255,0.1)',
                        color: 'white',
                        fontSize: '13px'
                    }}
                />
            </div>

            <div className="tree-container" id="tree-root">

                {isEditMode && (
                    <button
                        className="btn-primary"
                        onClick={onAddGrade}
                        style={{ width: '100%', marginBottom: '10px', padding: '5px', fontSize: '12px' }}
                    >
                        + Add Grade Level
                    </button>
                )}
                <ul>{renderTree(structure)}</ul>
            </div>


            <div style={{ padding: '15px' }}>
                {userRole === 'admin' && (
                    <div style={{ display: 'flex', gap: '5px', marginBottom: '10px' }}>
                        <button
                            className="btn-primary"
                            onClick={onUserManagement}
                            style={{ flex: 1 }}
                        >
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                            Manage Users
                        </button>
                        {/* Sync Inbox Bell */}
                        <button
                            onClick={onOpenSyncInbox}
                            title={`Sync Inbox${pendingSyncCount > 0 ? ` (${pendingSyncCount} pending)` : ''}`}
                            style={{
                                position: 'relative',
                                padding: '8px 12px',
                                borderRadius: '6px',
                                border: 'none',
                                background: pendingSyncCount > 0
                                    ? 'linear-gradient(135deg, #8b5cf6, #7c3aed)'
                                    : 'rgba(255,255,255,0.08)',
                                color: 'white',
                                cursor: 'pointer',
                                flexShrink: 0,
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center'
                            }}
                        >
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path>
                                <path d="M13.73 21a2 2 0 0 1-3.46 0"></path>
                            </svg>
                            {pendingSyncCount > 0 && (
                                <span style={{
                                    position: 'absolute',
                                    top: '-4px',
                                    right: '-4px',
                                    background: '#ef4444',
                                    color: 'white',
                                    borderRadius: '50%',
                                    minWidth: '18px',
                                    height: '18px',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    fontSize: '11px',
                                    fontWeight: 'bold',
                                    animation: 'pulse 1.5s infinite'
                                }}>
                                    {pendingSyncCount}
                                </span>
                            )}
                        </button>
                    </div>
                )}
                <button className="btn-primary" onClick={onShare}>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><circle cx="18" cy="5" r="3"></circle><circle cx="6" cy="12" r="3"></circle><circle cx="18" cy="19" r="3"></circle><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"></line><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"></line></svg>
                    Share / Export
                </button>
                <button
                    className="btn-primary"
                    onClick={onLogout}
                    style={{ marginTop: '10px', background: '#dc3545' }}
                >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px' }}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path><polyline points="16 17 21 12 16 7"></polyline><line x1="21" y1="12" x2="9" y2="12"></line></svg>
                    Logout
                </button>
            </div>
        </aside>
    );
}

export default Sidebar;
