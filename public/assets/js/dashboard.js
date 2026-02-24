const { ipcRenderer } = require('electron');
const fs = require('fs');

window.onerror = function (msg, url, lineNo, columnNo, error) {
    alert('Global Error: ' + msg + '\nLine: ' + lineNo);
    return false;
};

let currentStudentId = null;
let currentStudentData = null;
let saveTimer = null;

const editorContainer = document.getElementById('editor-content');
const statusLabel = document.getElementById('status-label');
const exportBtn = document.getElementById('export-btn');
const logoutBtn = document.getElementById('logout-btn');
const userManagementBtn = document.getElementById('user-management-btn');

document.addEventListener('DOMContentLoaded', async () => {
    try {
        if (!ipcRenderer) throw new Error('ipcRenderer is not defined');

        console.log('Dashboard initializing...');
        await loadSidebar();
        console.log('Sidebar loaded.');

        // Role-Based Access Control
        const role = await ipcRenderer.invoke('get-current-user-role');
        console.log('Role loaded:', role);
        if (role !== 'admin') {
            if (userManagementBtn) userManagementBtn.style.display = 'none';
        }

        if (exportBtn) exportBtn.addEventListener('click', shareToUSB); // Legacy button if exists
        if (logoutBtn) logoutBtn.addEventListener('click', () => ipcRenderer.send('logout'));
        if (userManagementBtn) userManagementBtn.addEventListener('click', showUserManagement);
    } catch (e) {
        console.error('Dashboard Initialization Error:', e);
        alert('Critical Error: ' + e.message + '\n' + e.stack); // Use native alert for critical startup errors
    }
});

// Show user management section
function showUserManagement() {
    console.log('showUserManagement called');
    // Hide welcome screen and any student records
    const welcomeScreen = document.getElementById('welcome-screen');
    const userManagementSection = document.getElementById('user-management-section');

    console.log('User Management Section:', userManagementSection);

    // Hide all pages (student records)
    const pages = editorContainer.querySelectorAll('.paper');
    pages.forEach(page => page.style.display = 'none');

    if (welcomeScreen) welcomeScreen.style.display = 'none';
    if (userManagementSection) {
        userManagementSection.style.display = 'block';
        userManagementSection.style.visibility = 'visible';
        console.log('User management section should be visible now');
    } else {
        console.error('User management section not found!');
    }

    // Deselect any active student
    document.querySelectorAll('.student-node').forEach(n => n.classList.remove('active'));
    currentStudentId = null;
    currentStudentData = null;
    statusLabel.textContent = 'User Management';
}


let isEditMode = false;
let currentStructure = null;

async function loadSidebar() {
    currentStructure = await ipcRenderer.invoke('get-structure');
    renderTree();
}

function renderTree() {
    const treeRoot = document.getElementById('tree-root');
    treeRoot.innerHTML = '';

    // Add "Add Grade" button if in edit mode
    if (isEditMode) {
        const addGradeBtn = document.createElement('button');
        addGradeBtn.className = 'btn-primary';
        addGradeBtn.style.width = '100%';
        addGradeBtn.style.marginBottom = '10px';
        addGradeBtn.style.padding = '5px';
        addGradeBtn.style.fontSize = '12px';
        addGradeBtn.textContent = '+ Add Grade Level';
        addGradeBtn.onclick = addGrade;
        treeRoot.appendChild(addGradeBtn);
    }

    const buildTree = (obj, parent, path = []) => {
        for (const key in obj) {
            const currentPath = [...path, key];
            const li = document.createElement('li');
            li.setAttribute('data-path', currentPath.join('|')); // For search identification

            // Container for the node content
            const nodeContent = document.createElement('div');
            nodeContent.style.display = 'flex';
            nodeContent.style.alignItems = 'center';
            nodeContent.style.justifyContent = 'space-between';

            if (Array.isArray(obj[key])) {
                // Leaf Node (Section -> Students)
                const details = document.createElement('details');
                details.setAttribute('data-node-path', currentPath.join('|'));
                const summary = document.createElement('summary');

                // Summary Content
                const summaryText = document.createElement('span');
                summaryText.textContent = key;
                summary.appendChild(summaryText);

                // Edit Controls for Section
                if (isEditMode) {
                    const controls = document.createElement('span');
                    controls.style.marginLeft = '10px';
                    controls.style.display = 'inline-flex';
                    controls.style.gap = '5px';

                    const addBtn = document.createElement('button');
                    addBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="8" cy="8" r="7" fill="#28a745" stroke="#ffffff" stroke-width="1"/>
                        <path d="M8 4V12M4 8H12" stroke="white" stroke-width="2" stroke-linecap="round"/>
                    </svg>`;
                    addBtn.title = 'Add Student';
                    addBtn.style.cursor = 'pointer';
                    addBtn.style.background = 'none';
                    addBtn.style.border = 'none';
                    addBtn.style.padding = '2px';
                    addBtn.style.display = 'inline-flex';
                    addBtn.style.alignItems = 'center';
                    addBtn.onclick = (e) => { e.preventDefault(); addStudent(currentPath); };

                    const delBtn = document.createElement('button');
                    delBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="8" cy="8" r="7" fill="#dc3545" stroke="#ffffff" stroke-width="1"/>
                        <path d="M5 5L11 11M11 5L5 11" stroke="white" stroke-width="2" stroke-linecap="round"/>
                    </svg>`;
                    delBtn.title = 'Delete Section';
                    delBtn.style.cursor = 'pointer';
                    delBtn.style.background = 'none';
                    delBtn.style.border = 'none';
                    delBtn.style.padding = '2px';
                    delBtn.style.display = 'inline-flex';
                    delBtn.style.alignItems = 'center';
                    delBtn.onclick = (e) => { e.preventDefault(); deleteItem(currentPath); };

                    controls.appendChild(addBtn);
                    controls.appendChild(delBtn);
                    summary.appendChild(controls);
                }

                details.appendChild(summary);

                const ul = document.createElement('ul');
                obj[key].forEach((student, index) => {
                    const sLi = document.createElement('li');
                    sLi.className = 'student-node';
                    sLi.setAttribute('data-student-name', student.name.toLowerCase());
                    sLi.setAttribute('data-student-path', [...currentPath, index].join('|'));

                    const sContent = document.createElement('div');
                    sContent.style.display = 'flex';
                    sContent.style.justifyContent = 'space-between';
                    sContent.style.alignItems = 'center';

                    const sName = document.createElement('span');
                    sName.textContent = student.name;
                    sContent.appendChild(sName);

                    if (isEditMode) {
                        const sDel = document.createElement('button');
                        sDel.innerHTML = `<svg width="14" height="14" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <circle cx="8" cy="8" r="7" fill="#dc3545" stroke="#ffffff" stroke-width="1"/>
                            <path d="M5 5L11 11M11 5L5 11" stroke="white" stroke-width="2" stroke-linecap="round"/>
                        </svg>`;
                        sDel.style.cursor = 'pointer';
                        sDel.style.background = 'none';
                        sDel.style.border = 'none';
                        sDel.style.padding = '2px';
                        sDel.style.marginLeft = '10px';
                        sDel.style.display = 'inline-flex';
                        sDel.style.alignItems = 'center';
                        sDel.onclick = (e) => {
                            e.stopPropagation();
                            deleteItem([...currentPath, index]); // Use index for array items
                        };
                        sContent.appendChild(sDel);
                    }

                    sLi.appendChild(sContent);
                    sLi.onclick = () => {
                        document.querySelectorAll('.student-node').forEach(n => n.classList.remove('active'));
                        sLi.classList.add('active');
                        loadStudent(student.id, student.name);
                    };
                    ul.appendChild(sLi);
                });
                details.appendChild(ul);
                li.appendChild(details);
            } else {
                // Branch Node (Grade/Strand)
                const details = document.createElement('details');
                details.setAttribute('data-node-path', currentPath.join('|'));
                const summary = document.createElement('summary');

                const summaryText = document.createElement('span');
                summaryText.textContent = key;
                summary.appendChild(summaryText);

                if (isEditMode) {
                    const controls = document.createElement('span');
                    controls.style.marginLeft = '10px';
                    controls.style.display = 'inline-flex';
                    controls.style.gap = '5px';

                    const addBtn = document.createElement('button');
                    addBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="8" cy="8" r="7" fill="#28a745" stroke="#ffffff" stroke-width="1"/>
                        <path d="M8 4V12M4 8H12" stroke="white" stroke-width="2" stroke-linecap="round"/>
                    </svg>`;
                    addBtn.title = 'Add Sub-item';
                    addBtn.style.cursor = 'pointer';
                    addBtn.style.background = 'none';
                    addBtn.style.border = 'none';
                    addBtn.style.padding = '2px';
                    addBtn.style.display = 'inline-flex';
                    addBtn.style.alignItems = 'center';
                    addBtn.onclick = (e) => {
                        e.preventDefault();
                        // Determine what to add based on depth
                        if (path.length === 0) addStrand(currentPath); // Grade -> Add Strand
                        else if (path.length === 1) addSection(currentPath); // Strand -> Add Section
                    };

                    const delBtn = document.createElement('button');
                    delBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="8" cy="8" r="7" fill="#dc3545" stroke="#ffffff" stroke-width="1"/>
                        <path d="M5 5L11 11M11 5L5 11" stroke="white" stroke-width="2" stroke-linecap="round"/>
                    </svg>`;
                    delBtn.title = 'Delete Item';
                    delBtn.style.cursor = 'pointer';
                    delBtn.style.background = 'none';
                    delBtn.style.border = 'none';
                    delBtn.style.padding = '2px';
                    delBtn.style.display = 'inline-flex';
                    delBtn.style.alignItems = 'center';
                    delBtn.onclick = (e) => { e.preventDefault(); deleteItem(currentPath); };

                    controls.appendChild(addBtn);
                    controls.appendChild(delBtn);
                    summary.appendChild(controls);
                }

                details.appendChild(summary);
                const ul = document.createElement('ul');
                buildTree(obj[key], ul, currentPath);
                details.appendChild(ul);
                li.appendChild(details);
            }
            parent.appendChild(li);
        }
    };

    const rootUl = document.createElement('ul');
    if (currentStructure) buildTree(currentStructure, rootUl);
    treeRoot.appendChild(rootUl);
}

// Search Logic
const searchInput = document.getElementById('sidebar-search');
if (searchInput) {
    let lastFoundStudent = null;

    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (!query) {
            // Collapse all if empty? Or just remove highlights
            document.querySelectorAll('.search-highlight').forEach(el => el.classList.remove('search-highlight'));
            lastFoundStudent = null;
            return;
        }

        // Search for students
        const studentNodes = document.querySelectorAll('.student-node');
        let found = false;

        // Clear previous highlights
        document.querySelectorAll('.search-highlight').forEach(el => el.classList.remove('search-highlight'));
        lastFoundStudent = null;

        for (const node of studentNodes) {
            const name = node.getAttribute('data-student-name');
            if (name && name.includes(query)) {
                // Found a match
                found = true;
                lastFoundStudent = node;

                // Highlight
                node.classList.add('search-highlight');

                // Expand parents
                let parent = node.parentElement;
                while (parent) {
                    if (parent.tagName === 'DETAILS') {
                        parent.open = true;
                    }
                    parent = parent.parentElement;
                }

                // Scroll into view
                node.scrollIntoView({ behavior: 'smooth', block: 'center' });
                break; // Jump to first match
            }
        }

        if (!found) {
            // Search for sections/strands/grades
            const details = document.querySelectorAll('details[data-node-path]');
            for (const d of details) {
                const path = d.getAttribute('data-node-path');
                // Check if the key (last part of path) matches
                const parts = path.split('|');
                const key = parts[parts.length - 1].toLowerCase();

                if (key.includes(query)) {
                    d.open = true;
                    d.querySelector('summary').classList.add('search-highlight');

                    // Expand parents
                    let parent = d.parentElement;
                    while (parent) {
                        if (parent.tagName === 'DETAILS') {
                            parent.open = true;
                        }
                        parent = parent.parentElement;
                    }

                    d.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    break;
                }
            }
        }
    });

    // Handle Enter key to auto-open student
    searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && lastFoundStudent) {
            lastFoundStudent.click();
        }
    });
}

// Listeners
document.getElementById('edit-structure-btn')?.addEventListener('click', () => {
    isEditMode = !isEditMode;
    const btn = document.getElementById('edit-structure-btn');
    btn.textContent = isEditMode ? '✅ Done Editing' : '✏️ Edit Structure';
    btn.style.background = isEditMode ? '#28a745' : ''; // Green when editing
    renderTree();
});

document.getElementById('refresh-tree-btn')?.addEventListener('click', loadSidebar);

async function loadStudent(id, name) {
    currentStudentId = id;

    // Hide welcome screen and user management
    const welcomeScreen = document.getElementById('welcome-screen');
    const userManagementSection = document.getElementById('user-management-section');
    if (welcomeScreen) welcomeScreen.style.display = 'none';
    if (userManagementSection) userManagementSection.style.display = 'none';

    // Remove existing papers
    const existingPages = editorContainer.querySelectorAll('.paper');
    existingPages.forEach(p => p.remove());

    // Show loading
    let loader = document.getElementById('student-loader');
    if (!loader) {
        loader = document.createElement('div');
        loader.id = 'student-loader';
        loader.style.textAlign = 'center';
        loader.style.padding = '50px';
        loader.textContent = 'Loading...';
        editorContainer.appendChild(loader);
    }
    loader.style.display = 'block';

    let data = await ipcRenderer.invoke('load-student', id);
    if (!data) data = generateEmptyRecord(id, name);
    currentStudentData = data;

    // Hide loader before rendering
    loader.style.display = 'none';
    renderForm();
}

function generateEmptyRecord(id, name) {
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
        info: { lname, fname, mname: '', lrn: '', sex: '', birthdate: '', admissionDate: '' },
        eligibility: {
            hsCompleter: false, hsGenAve: '', jhsCompleter: false, jhsGenAve: '',
            gradDate: '', schoolName: '', schoolAddress: '',
            pept: false, peptRating: '', als: false, alsRating: '',
            examDate: '', clcName: '', othersSpec: ''
        },
        semester1: makeSem(),
        semester2: makeSem(),
        semester3: makeSem(),
        semester4: makeSem(),
        certification: { trackStrand: '', genAve: '', awards: '', gradDate: '', schoolHead: '', certDate: '', remarks: '', dateIssued: '' }
    };
}

function renderForm() {
    // Hide welcome/user-management/loader just in case
    const welcomeScreen = document.getElementById('welcome-screen');
    const userManagementSection = document.getElementById('user-management-section');
    const loader = document.getElementById('student-loader');

    if (welcomeScreen) welcomeScreen.style.display = 'none';
    if (userManagementSection) userManagementSection.style.display = 'none';
    if (loader) loader.style.display = 'none';

    // Remove old papers
    const existingPages = editorContainer.querySelectorAll('.paper');
    existingPages.forEach(p => p.remove());

    const d = currentStudentData;

    // Ensure all semesters exist (for backward compatibility)
    if (!d.semester3) d.semester3 = generateEmptyRecord().semester1;
    if (!d.semester4) d.semester4 = generateEmptyRecord().semester1;

    // === PAGE 1: Header + Learner Info + Eligibility + Semesters 1 & 2 ===
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
                <td style="width:33%; font-size:6pt; font-weight:bold;">LAST NAME: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 70px); font-weight:normal;"><input class="inp" value="${d.info.lname}" oninput="up('info','lname',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:34%; font-size:6pt; font-weight:bold;">FIRST NAME: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 75px); font-weight:normal;"><input class="inp" value="${d.info.fname}" oninput="up('info','fname',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:33%; font-size:6pt; font-weight:bold;">MIDDLE NAME: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 90px); font-weight:normal;"><input class="inp" value="${d.info.mname}" oninput="up('info','mname',this.value)" style="width:100%; border:none;"></span></td>
            </tr>
            <tr>
                <td style="font-size:6pt; font-weight:bold;">LRN: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 35px); font-weight:normal;"><input class="inp" value="${d.info.lrn}" oninput="up('info','lrn',this.value)" style="width:100%; border:none;"></span></td>
                <td colspan="2" style="font-size:6pt; font-weight:bold;">Date of Birth (mm/dd/yyyy): <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 145px); font-weight:normal;"><input class="inp" value="${d.info.birthdate}" oninput="up('info','birthdate',this.value)" style="width:100%; border:none;"></span></td>
            </tr>
            <tr>
                <td style="font-size:6pt; font-weight:bold;">Sex: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 30px); font-weight:normal;"><input class="inp" value="${d.info.sex}" oninput="up('info','sex',this.value)" style="width:100%; border:none;"></span></td>
                <td colspan="2" style="font-size:6pt; font-weight:bold;">Date of SHS Admission (mm/dd/yyyy): <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 195px); font-weight:normal;"><input class="inp" value="${d.info.admissionDate}" oninput="up('info','admissionDate',this.value)" style="width:100%; border:none;"></span></td>
            </tr>
        </table>

        <div class="hdr">ELIGIBILITY FOR SHS ENROLMENT</div>
        <table style="width:100%; border:none; border-collapse:collapse; font-size:5.5pt;">
            <tr>
                <td style="width:18%; border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.hsCompleter ? 'checked' : ''} onchange="up('eligibility','hsCompleter',this.checked)"> High School Completer*</td>
                <td style="width:14%; border:none; padding:1px 2px; white-space:nowrap;">Gen. Ave: <input class="inp" value="${d.eligibility.hsGenAve}" oninput="up('eligibility','hsGenAve',this.value)" style="border:none; border-bottom:1px solid #000; width:50%; display:inline-block;"></td>
                <td style="width:18%; border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.jhsCompleter ? 'checked' : ''} onchange="up('eligibility','jhsCompleter',this.checked)"> JHS Completer</td>
                <td style="width:14%; border:none; padding:1px 2px; white-space:nowrap;">Gen. Ave: <input class="inp" value="${d.eligibility.jhsGenAve}" oninput="up('eligibility','jhsGenAve',this.value)" style="border:none; border-bottom:1px solid #000; width:50%; display:inline-block;"></td>
                <td style="width:36%; border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.others ? 'checked' : ''} onchange="up('eligibility','others',this.checked)"> Others: <input class="inp" value="${d.eligibility.othersSpec}" oninput="up('eligibility','othersSpec',this.value)" style="border:none; border-bottom:1px solid #000; width:calc(100% - 65px); display:inline-block;"></td>
            </tr>
            <tr>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">Date of Graduation: <input class="inp" value="${d.eligibility.gradDate}" oninput="up('eligibility','gradDate',this.value)" style="border:none; border-bottom:1px solid #000; width:55%; display:inline-block;"></td>
                <td style="border:none; padding:1px 2px; white-space:nowrap;">Name of School: <input class="inp" value="${d.eligibility.schoolName}" oninput="up('eligibility','schoolName',this.value)" style="border:none; border-bottom:1px solid #000; width:calc(100% - 90px); display:inline-block;"></td>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">School Address: <input class="inp" value="${d.eligibility.schoolAddress}" oninput="up('eligibility','schoolAddress',this.value)" style="border:none; border-bottom:1px solid #000; width:calc(100% - 100px); display:inline-block;"></td>
            </tr>
            <tr>
                <td style="border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.pept ? 'checked' : ''} onchange="up('eligibility','pept',this.checked)"> PEPT Passer**</td>
                <td style="border:none; padding:1px 2px; white-space:nowrap;">Rating: <input class="inp" value="${d.eligibility.peptRating}" oninput="up('eligibility','peptRating',this.value)" style="border:none; border-bottom:1px solid #000; width:60%; display:inline-block;"></td>
                <td style="border:none; padding:1px 2px; white-space:nowrap;"><input type="checkbox" ${d.eligibility.als ? 'checked' : ''} onchange="up('eligibility','als',this.checked)"> ALS A&E***</td>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">Rating: <input class="inp" value="${d.eligibility.alsRating}" oninput="up('eligibility','alsRating',this.value)" style="border:none; border-bottom:1px solid #000; width:calc(100% - 55px); display:inline-block;"></td>
            </tr>
            <tr>
                <td colspan="2" style="border:none; padding:1px 2px; white-space:nowrap;">Date of Exam: <input class="inp" value="${d.eligibility.examDate}" oninput="up('eligibility','examDate',this.value)" style="border:none; border-bottom:1px solid #000; width:55%; display:inline-block;"></td>
                <td colspan="3" style="border:none; padding:1px 2px; white-space:nowrap;">CLC Name and Address: <input class="inp" value="${d.eligibility.clcName}" oninput="up('eligibility','clcName',this.value)" style="border:none; border-bottom:1px solid #000; width:calc(100% - 135px); display:inline-block;"></td>
            </tr>
            <tr style="font-size:5pt; font-style:italic;">
                <td colspan="2" style="border:none; padding:1px 2px;">*HS Completers graduated under old curriculum</td>
                <td colspan="3" style="border:none; padding:1px 2px;">***ALS A&E - Alternative Learning System Accred. & Equiv. Test</td>
            </tr>
            <tr style="font-size:5pt; font-style:italic;">
                <td colspan="5" style="border:none; padding:1px 2px;">**PEPT - Philippine Educational Placement Test</td>
            </tr>
        </table>

        ${renderSemester(d.semester1, 'semester1')}
        ${renderSemester(d.semester2, 'semester2')}
    `;
    editorContainer.appendChild(page1);

    // === PAGE 2: Semesters 3 & 4 + Certification ===
    const page2 = document.createElement('div');
    page2.className = 'paper';
    page2.innerHTML = `
        <div style="text-align:right; font-size:6pt; margin-bottom:3px;">Page 2 &nbsp; Form 137-SHS</div>
        ${renderSemester(d.semester3, 'semester3')}
        ${renderSemester(d.semester4, 'semester4')}
        
        <div style="font-size:6pt; margin-top:3px;">
            <div style="margin-bottom:2px;">
                <span style="font-weight:bold;">Track/Strand Accomplished:</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:340px; margin:0 20px 0 5px; vertical-align:bottom;">
                    <input class="inp" value="${d.certification.trackStrand}" oninput="up('certification','trackStrand',this.value)" style="width:100%; border:none; background:transparent; padding:0;">
                </span>
                <span style="font-weight:bold;">SHS General Average:</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:120px; margin-left:5px; vertical-align:bottom;">
                    <input class="inp" value="${d.certification.genAve}" oninput="up('certification','genAve',this.value)" style="width:100%; border:none; background:transparent; padding:0;">
                </span>
            </div>
            <div>
                <span style="font-weight:bold;">Awards/Honors Received:</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:318px; margin:0 20px 0 5px; vertical-align:bottom;">
                    <input class="inp" value="${d.certification.awards}" oninput="up('certification','awards',this.value)" style="width:100%; border:none; background:transparent; padding:0;">
                </span>
                <span style="font-weight:bold;">Date of SHS Graduation (MM/DD/YYYY):</span>
                <span style="border-bottom:1px solid #000; display:inline-block; width:90px; margin-left:5px; vertical-align:bottom;">
                    <input type="date" class="inp" value="${d.certification.gradDate}" oninput="up('certification','gradDate',this.value)" style="width:100%; border:none; background:transparent; padding:0;">
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
                        <input type="date" class="inp" value="${d.certification.certDate}" oninput="up('certification','certDate',this.value)" style="width:100%; border:none; background:transparent; text-align:center; padding:0; font-size:5.5pt;">
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
                    <textarea class="inp" style="width:100%; border:none; background:transparent; font-family:inherit; resize:none; min-height:40px; font-size:6pt; padding:0; line-height:1.3; margin-bottom:5px;" oninput="up('certification','remarks',this.value)">${d.certification.remarks}</textarea>
                    <div style="margin-top:5px;">
                        <span style="font-weight:bold; margin-right:5px;">Date Issued (MM/DD/YYYY):</span>
                        <span style="border-bottom:1px solid #000; display:inline-block; width:120px; vertical-align:bottom;">
                            <input type="date" class="inp" value="${d.certification.dateIssued}" oninput="up('certification','dateIssued',this.value)" style="width:100%; border:none; background:transparent; padding:0; font-size:6pt;">
                        </span>
                    </div>
                </td>
                <td style="border:none; border-left:1px solid #000; border-bottom:1px solid #000;"></td>
            </tr>
        </table>
    `;
    editorContainer.appendChild(page2);
}

function renderSemester(sem, key) {
    return `
        <div class="hdr" style="margin-top:5px;">SCHOLASTIC RECORD</div>
        <table class="clean-tbl" style="width:100%; font-size:6pt;">
            <tr>
                <td style="width:35%; font-weight:bold;">SCHOOL: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 55px); font-weight:normal;"><input class="inp" value="${sem.school}" oninput="upSem('${key}','school',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:15%; font-weight:bold;">SCHOOL ID: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 75px); font-weight:normal;"><input class="inp" value="${sem.schoolId}" oninput="upSem('${key}','schoolId',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:15%; font-weight:bold;">GRADE LEVEL: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 90px); font-weight:normal;"><input class="inp" value="${sem.gradeLevel}" oninput="upSem('${key}','gradeLevel',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:15%; font-weight:bold;">SY: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 25px); font-weight:normal;"><input class="inp" value="${sem.sy}" oninput="upSem('${key}','sy',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:20%; font-weight:bold;">SEM: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 35px); font-weight:normal;"><input class="inp" value="${sem.sem}" oninput="upSem('${key}','sem',this.value)" style="width:100%; border:none;"></span></td>
            </tr>
            <tr>
                <td colspan="3" style="font-weight:bold;">TRACK/STRAND: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 110px); font-weight:normal;"><input class="inp" value="${sem.trackStrand}" oninput="upSem('${key}','trackStrand',this.value)" style="width:100%; border:none;"></span></td>
                <td colspan="2" style="font-weight:bold;">SECTION: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 60px); font-weight:normal;"><input class="inp" value="${sem.section}" oninput="upSem('${key}','section',this.value)" style="width:100%; border:none;"></span></td>
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
                    <th width="7%"><input class="inp c" style="font-weight:bold; background:transparent; border:none; padding:0;" value="${sem.q1Label || '1ST'}" oninput="upSem('${key}','q1Label',this.value)"></th>
                    <th width="7%"><input class="inp c" style="font-weight:bold; background:transparent; border:none; padding:0;" value="${sem.q2Label || '2ND'}" oninput="upSem('${key}','q2Label',this.value)"></th>
                </tr>
            </thead>
            <tbody>
                ${sem.subjects.slice(0, 9).map((s, i) => `
                    <tr>
                        <td>
                            <select class="inp c" style="font-size:5pt; padding:0;" onchange="upSub('${key}',${i},'type',this.value)">
                                <option value="" ${s.type === '' ? 'selected' : ''}></option>
                                <option value="Core" ${s.type === 'Core' ? 'selected' : ''}>Core</option>
                                <option value="Applied" ${s.type === 'Applied' ? 'selected' : ''}>Applied</option>
                                <option value="Specialized" ${s.type === 'Specialized' ? 'selected' : ''}>Specialized</option>
                            </select>
                        </td>
                        <td><input class="inp" value="${s.subject}" oninput="upSub('${key}',${i},'subject',this.value)"></td>
                        <td><input class="inp c" value="${s.q1}" oninput="upSub('${key}',${i},'q1',this.value); calcGrades('${key}', ${i});"></td>
                        <td><input class="inp c" value="${s.q2}" oninput="upSub('${key}',${i},'q2',this.value); calcGrades('${key}', ${i});"></td>
                        <td><input class="inp c" value="${s.final}" id="final-${key}-${i}" oninput="upSub('${key}',${i},'final',this.value); calcGenAve('${key}');"></td>
                        <td><input class="inp c" value="${s.action}" id="action-${key}-${i}" oninput="upSub('${key}',${i},'action',this.value)"></td>
                    </tr>
                `).join('')}
                <tr>
                    <td colspan="4" style="text-align:right; font-weight:bold; padding-right:5px;">General Ave. for the Semester:</td>
                    <td><input class="inp c" value="${sem.genAve}" oninput="upSem('${key}','genAve',this.value)" id="genAve-${key}" style="border:none; width:100%;"></td>
                    <td></td>
                </tr>
            </tbody>
        </table>

        <div style="font-size:6pt; font-weight:bold; margin-top:2px; display:flex; align-items:flex-end;">
            <span style="margin-right:5px;">REMARKS:</span>
            <div style="flex-grow:1; border-bottom:1px solid #000;">
                <input class="inp" value="${sem.remarks}" oninput="upSem('${key}','remarks',this.value)" style="width:100%; border:none; background:transparent;">
            </div>
        </div>

        <table class="clean-tbl" style="font-size:5.5pt; width:100%;">
            <tr>
                <td style="width:33%;">Prepared by: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 80px);"><input class="inp" value="${sem.adviserName}" oninput="upSem('${key}','adviserName',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:34%;">Certified: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 65px);"><input class="inp" value="${sem.certName}" oninput="upSem('${key}','certName',this.value)" style="width:100%; border:none;"></span></td>
                <td style="width:33%;">Date Checked: <span style="border-bottom:1px solid #000; display:inline-block; width:calc(100% - 90px);"><input class="inp" value="${sem.dateChecked}" oninput="upSem('${key}','dateChecked',this.value)" style="width:100%; border:none;"></span></td>
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
                <td colspan="2">Conducted from: <input class="inp" value="${sem.remedial.from}" oninput="upRem('${key}','from',this.value)" style="border-bottom:1px solid #000; width:50px; display:inline-block; border:none;"></td>
                <td>to: <input class="inp" value="${sem.remedial.to}" oninput="upRem('${key}','to',this.value)" style="border-bottom:1px solid #000; width:50px; display:inline-block; border:none;"></td>
                <td style="font-weight:bold;">SCHOOL: <input class="inp" value="${sem.remedial.school}" oninput="upRem('${key}','school',this.value)" style="border-bottom:1px solid #000; width:calc(100% - 60px); display:inline-block; border:none;"></td>
                <td colspan="2" style="font-weight:bold;">SCHOOL ID: <input class="inp" value="${sem.remedial.schoolId}" oninput="upRem('${key}','schoolId',this.value)" style="border-bottom:1px solid #000; width:calc(100% - 75px); display:inline-block; border:none;"></td>
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
                ${sem.remedial.subjects.slice(0, 4).map((rs, ri) => `
                    <tr>
                        <td><input class="inp c" value="${rs.type}" oninput="upRemSub('${key}',${ri},'type',this.value)"></td>
                        <td><input class="inp" value="${rs.subject}" oninput="upRemSub('${key}',${ri},'subject',this.value)"></td>
                        <td><input class="inp c" value="${rs.semGrade}" oninput="upRemSub('${key}',${ri},'semGrade',this.value)"></td>
                        <td><input class="inp c" value="${rs.remedialMark}" oninput="upRemSub('${key}',${ri},'remedialMark',this.value)"></td>
                        <td><input class="inp c" value="${rs.recomputedGrade}" oninput="upRemSub('${key}',${ri},'recomputedGrade',this.value)"></td>
                        <td><input class="inp c" value="${rs.action}" oninput="upRemSub('${key}',${ri},'action',this.value)"></td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
        <table class="form" style="font-size:5.5pt;">
            <tr><td width="50%">Name of Teacher/Adviser: <input class="inp" value="${sem.remedial.teacherName}" oninput="upRem('${key}','teacherName',this.value)" style="border-bottom:1px solid #000; width:calc(100% - 145px);"></td><td>Signature:</td></tr>
        </table>
    `;
}

function up(section, field, val) {
    currentStudentData[section][field] = val;
    save();
}
function upSem(semKey, field, val) {
    currentStudentData[semKey][field] = val;
    save();
}
function upSub(semKey, idx, field, val) {
    currentStudentData[semKey].subjects[idx][field] = val;
    save();
}
function upRem(semKey, field, val) {
    currentStudentData[semKey].remedial[field] = val;
    save();
}
function upRemSub(semKey, idx, field, val) {
    currentStudentData[semKey].remedial.subjects[idx][field] = val;
    save();
}

function save() {
    statusLabel.innerText = "Unsaved...";
    statusLabel.style.color = "orange";
    clearTimeout(saveTimer);
    saveTimer = setTimeout(async () => {
        statusLabel.innerText = "Saving...";
        await ipcRenderer.invoke('save-student', { id: currentStudentId, data: currentStudentData });
        statusLabel.innerText = "Saved";
        statusLabel.style.color = "green";
    }, 1000);
}

async function shareToUSB() {
    const drives = await ipcRenderer.invoke('scan-usb');
    if (!drives.length) {
        await customAlert("No USB found.");
        return;
    }
    const drive = drives[0];
    if (!confirm(`Export to ${drive.label}?`)) return;
    const content = `<html><head><title>${currentStudentData.info.lname} - Form 137</title>
    <style>${fs.readFileSync('assets/css/form137.css', 'utf8')}</style></head>
    <body>${editorContainer.innerHTML}</body></html>`;
    const result = await ipcRenderer.invoke('export-file', {
        path: drive.path,
        filename: `Form137_${currentStudentData.info.lname}.html`,
        content
    });
    await customAlert(result.success ? "Export successful!" : "Export failed: " + result.error);
}

window.up = up;
window.upSem = upSem;
window.upSub = upSub;
window.upRem = upRem;
window.upRemSub = upRemSub;

// --- STRUCTURE MANAGEMENT HELPERS ---

function saveStructure() {
    ipcRenderer.send('update-structure', currentStructure);
    loadSidebar(); // Refresh tree
}

// Custom prompt replacement (Electron doesn't support native prompt())
function customPrompt(title) {
    return new Promise((resolve) => {
        const modal = document.getElementById('custom-prompt-modal');
        const titleEl = document.getElementById('custom-prompt-title');
        const inputEl = document.getElementById('custom-prompt-input');
        const okBtn = document.getElementById('custom-prompt-ok');
        const cancelBtn = document.getElementById('custom-prompt-cancel');

        titleEl.textContent = title;
        inputEl.value = '';
        modal.style.display = 'flex';

        // Use setTimeout to ensure modal is visible before focusing
        setTimeout(() => inputEl.focus(), 50);

        let resolved = false;

        const cleanup = (value) => {
            if (resolved) return; // Prevent double resolution
            resolved = true;

            modal.style.display = 'none';
            okBtn.removeEventListener('click', handleOk);
            cancelBtn.removeEventListener('click', handleCancel);
            inputEl.removeEventListener('keydown', handleKeydown);
            resolve(value);
        };

        const handleOk = () => cleanup(inputEl.value.trim() || null);
        const handleCancel = () => cleanup(null);
        const handleKeydown = (e) => {
            if (e.key === 'Enter') handleOk();
            if (e.key === 'Escape') handleCancel();
        };

        okBtn.addEventListener('click', handleOk);
        cancelBtn.addEventListener('click', handleCancel);
        inputEl.addEventListener('keydown', handleKeydown);
    });
}

// Custom alert replacement
function customAlert(message) {
    return new Promise((resolve) => {
        const modal = document.getElementById('custom-alert-modal');
        const messageEl = document.getElementById('custom-alert-message');
        const okBtn = document.getElementById('custom-alert-ok');

        messageEl.textContent = message;
        modal.style.display = 'flex';

        // Focus OK button for keyboard accessibility
        setTimeout(() => okBtn.focus(), 50);

        const cleanup = () => {
            modal.style.display = 'none';
            okBtn.removeEventListener('click', handleOk);
            document.removeEventListener('keydown', handleKeydown);
            resolve();
        };

        const handleOk = () => cleanup();
        const handleKeydown = (e) => {
            if (e.key === 'Enter' || e.key === 'Escape') handleOk();
        };

        okBtn.addEventListener('click', handleOk);
        document.addEventListener('keydown', handleKeydown);
    });
}

async function addGrade() {
    const name = await customPrompt("Enter Grade Level Name (e.g., Grade 11):");
    if (!name) return;
    if (currentStructure[name]) {
        await customAlert("Grade Level already exists.");
        return;
    }
    currentStructure[name] = {};
    saveStructure();
}

async function addStrand(path) {
    const name = await customPrompt("Enter Strand Name (e.g., TVL - ICT):");
    if (!name) return;
    if (currentStructure[path[0]][name]) {
        await customAlert("Strand already exists.");
        return;
    }
    currentStructure[path[0]][name] = {};
    saveStructure();
}

async function addSection(path) {
    const name = await customPrompt("Enter Section Name (e.g., Section A):");
    if (!name) return;
    if (currentStructure[path[0]][path[1]][name]) {
        await customAlert("Section already exists.");
        return;
    }
    currentStructure[path[0]][path[1]][name] = [];
    saveStructure();
}

async function addStudent(path) {
    const name = await customPrompt("Enter Student Name (Last Name, First Name):");
    if (!name) return;
    const id = Date.now().toString(); // Simple ID generation
    currentStructure[path[0]][path[1]][path[2]].push({ id, name });
    saveStructure();
}

// Custom confirm replacement
function customConfirm(message) {
    return new Promise((resolve) => {
        const modal = document.getElementById('custom-confirm-modal');
        const messageEl = document.getElementById('custom-confirm-message');
        const yesBtn = document.getElementById('custom-confirm-yes');
        const noBtn = document.getElementById('custom-confirm-no');

        messageEl.textContent = message;
        modal.style.display = 'flex';

        // Focus No button for safety
        setTimeout(() => noBtn.focus(), 50);

        const cleanup = (result) => {
            modal.style.display = 'none';
            yesBtn.removeEventListener('click', handleYes);
            noBtn.removeEventListener('click', handleNo);
            document.removeEventListener('keydown', handleKeydown);
            resolve(result);
        };

        const handleYes = () => cleanup(true);
        const handleNo = () => cleanup(false);
        const handleKeydown = (e) => {
            if (e.key === 'Escape') handleNo();
        };

        yesBtn.addEventListener('click', handleYes);
        noBtn.addEventListener('click', handleNo);
        document.addEventListener('keydown', handleKeydown);
    });
}

async function deleteItem(path) {
    if (!await customConfirm("Are you sure you want to delete this item? This cannot be undone.")) return;

    if (path.length === 1) { // Delete Grade
        delete currentStructure[path[0]];
    } else if (path.length === 2) { // Delete Strand
        delete currentStructure[path[0]][path[1]];
    } else if (path.length === 3) { // Delete Section
        delete currentStructure[path[0]][path[1]][path[2]];
    } else if (path.length === 4) { // Delete Student
        // path[3] is the index
        currentStructure[path[0]][path[1]][path[2]].splice(path[3], 1);
    }
    saveStructure();
}

// --- SHARE FEATURE ---

const shareModal = document.getElementById('share-modal');
const shareBtn = document.getElementById('share-btn');
const closeShareBtn = document.getElementById('close-share-modal');
const shareUsbBtn = document.getElementById('share-usb-btn');
const shareFileBtn = document.getElementById('share-file-btn');
const importBtn = document.getElementById('import-btn');

if (shareBtn) shareBtn.addEventListener('click', async () => {
    if (!currentStudentData) {
        await customAlert("Please select a student first.");
        return;
    }
    shareModal.style.display = 'flex';
});

if (closeShareBtn) closeShareBtn.addEventListener('click', () => {
    shareModal.style.display = 'none';
});

if (shareUsbBtn) shareUsbBtn.addEventListener('click', exportToUSB);
if (shareFileBtn) shareFileBtn.addEventListener('click', exportToFile);
if (importBtn) importBtn.addEventListener('click', importStudent);

async function exportToUSB() {
    const drives = await ipcRenderer.invoke('scan-usb');
    if (!drives.length) {
        await customAlert("No USB found.");
        return;
    }
    const drive = drives[0];
    if (!confirm(`Export to ${drive.label}?`)) return;

    // 1. Export HTML (for printing)
    const htmlContent = `<html><head><title>${currentStudentData.info.lname} - Form 137</title>
    <style>${fs.readFileSync('assets/css/form137.css', 'utf8')}</style></head>
    <body>${editorContainer.innerHTML}</body></html>`;

    await ipcRenderer.invoke('export-file', {
        path: drive.path,
        filename: `Form137_${currentStudentData.info.lname}.html`,
        content: htmlContent
    });

    // 2. Export JSON (for importing)
    await ipcRenderer.invoke('export-file', {
        path: drive.path,
        filename: `StudentData_${currentStudentData.info.lname}.json`,
        content: JSON.stringify(currentStudentData, null, 2)
    });

    await customAlert("Export successful! Saved HTML and JSON to USB.");
    shareModal.style.display = 'none';
}

async function exportToFile() {
    const defaultName = `StudentData_${currentStudentData.info.lname}.json`;
    const filePath = await ipcRenderer.invoke('save-file-dialog', defaultName);
    if (!filePath) return;

    require('fs').writeFileSync(filePath, JSON.stringify(currentStudentData, null, 2));
    await customAlert("Export successful! You can now share this file via Bluetooth or Hotspot.");
    shareModal.style.display = 'none';
}

async function importStudent() {
    const filePath = await ipcRenderer.invoke('open-file-dialog');
    if (!filePath) return;

    const content = await ipcRenderer.invoke('read-file', filePath);
    if (!content) {
        await customAlert("Failed to read file.");
        return;
    }

    try {
        const studentData = JSON.parse(content);
        const name = `${studentData.info.lname}, ${studentData.info.fname}`;
        const id = studentData.id || Date.now().toString(); // Use existing ID or generate new

        // Determine Structure Placement (Grade -> Strand -> Section)
        // Try to find info in Semester 1, or default
        let grade = studentData.semester1.gradeLevel || "Unassigned Grade";
        let strand = studentData.semester1.trackStrand || "Unassigned Strand";
        let section = studentData.semester1.section || "Unassigned Section";

        // Auto-Create Structure
        if (!currentStructure[grade]) currentStructure[grade] = {};
        if (!currentStructure[grade][strand]) currentStructure[grade][strand] = {};
        if (!currentStructure[grade][strand][section]) currentStructure[grade][strand][section] = [];

        // Check if student already exists in this section
        const exists = currentStructure[grade][strand][section].find(s => s.name === name);
        if (exists) {
            if (!confirm(`Student ${name} already exists in ${grade} - ${section}. Overwrite?`)) return;
            // Update ID if overwriting? Or just keep existing?
            // Let's assume we update the record data but keep the structure entry
        } else {
            currentStructure[grade][strand][section].push({ id, name });
            saveStructure(); // Updates sidebar
        }

        // Save Student Record
        await ipcRenderer.invoke('save-student', { id, data: studentData });

        await customAlert(`Imported ${name} successfully!`);
        loadSidebar(); // Refresh tree
        loadStudent(id, name); // Load the imported student

    } catch (e) {
        await customAlert("Invalid Student Data File: " + e.message);
    }
}

// Print Button Handler
const printBtn = document.getElementById('print-record-btn');
if (printBtn) {
    printBtn.addEventListener('click', async () => {
        if (!currentStudentId || !currentStudentData) {
            await customAlert('Please select a student to print their record.');
            return;
        }
        // Save data for the print window to access
        localStorage.setItem('printStudentData', JSON.stringify(currentStudentData));

        // Open print preview in a new window
        const printWindow = window.open('print-preview.html', '_blank', 'width=1000,height=800,menubar=no,toolbar=no,location=no,status=no');

        if (!printWindow) {
            await customAlert('Pop-up blocked! Please allow pop-ups for this application.');
        }
    });
}

window.calcGrades = function (key, index) {
    // Find the semester data using the global currentStudentData
    if (!currentStudentData || !currentStudentData[key]) return;

    const sem = currentStudentData[key];
    const sub = sem.subjects[index];
    const q1 = parseFloat(sub.q1);
    const q2 = parseFloat(sub.q2);

    if (!isNaN(q1) && !isNaN(q2)) {
        const final = Math.round((q1 + q2) / 2);
        sub.final = final;
        sub.action = final >= 75 ? 'Passed' : 'Failed';

        // Update DOM
        const finalInput = document.getElementById(`final-${key}-${index}`);
        const actionInput = document.getElementById(`action-${key}-${index}`);
        if (finalInput) finalInput.value = final;
        if (actionInput) actionInput.value = sub.action;

        // Trigger save since we modified data directly
        if (window.save) window.save();

        calcGenAve(key);
    }
}

window.calcGenAve = function (key) {
    if (!currentStudentData || !currentStudentData[key]) return;

    const sem = currentStudentData[key];
    let total = 0;
    let count = 0;

    sem.subjects.forEach(s => {
        const val = parseFloat(s.final);
        if (!isNaN(val)) {
            total += val;
            count++;
        }
    });

    if (count > 0) {
        const ave = Math.round(total / count);
        sem.genAve = ave;

        // Update DOM
        const genAveInput = document.getElementById(`genAve-${key}`);
        if (genAveInput) genAveInput.value = ave;

        // Trigger save
        if (window.save) window.save();
    }
}
