import os
import sys
import json
import time
import datetime
import subprocess
from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
PORT = 8787
API_KEY = "sf10-bridge-2024"
TEMPLATE_NAME = 'PLACEHOLDER(ALL).xlsx'

# Robust path detection for PyInstaller
if hasattr(sys, '_MEIPASS'):
    EXE_DIR = os.path.dirname(sys.executable)
else:
    EXE_DIR = os.path.dirname(os.path.abspath(__file__))

PATHS_TO_CHECK = [
    os.path.join(EXE_DIR, TEMPLATE_NAME),
    os.path.join(EXE_DIR, 'release', TEMPLATE_NAME),
    os.path.join(os.getcwd(), 'release', TEMPLATE_NAME),
    os.path.join(os.getcwd(), TEMPLATE_NAME)
]

TEMPLATE_PATH = TEMPLATE_NAME
for p in PATHS_TO_CHECK:
    if os.path.exists(p):
        TEMPLATE_PATH = p
        break

OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "form137-exports")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

LOG_FILE = os.path.join(EXE_DIR, "excel-bridge-python.log")
def log_it(msg):
    try:
        with open(LOG_FILE, "a") as f:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{ts}] {msg}\n")
    except:
        pass

# --- HELPER: SAFE WRITE ---
def safe_write(ws, row_or_coord, col_or_val, val=None):
    if val is None:
        coord = str(row_or_coord).upper()
        write_val = col_or_val
    else:
        coord = f"{get_column_letter(col_or_val)}{row_or_coord}"
        write_val = val
    try:
        ws[coord] = write_val
    except AttributeError:
        # Handle merged cells
        for merged_range in ws.merged_cells.ranges:
            if coord in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = write_val
                return
        log_it(f"Failed write: {coord}")

app = Flask(__name__)
CORS(app)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"success": True, "status": "Running", "template": TEMPLATE_PATH})

@app.route('/open-excel', methods=['POST'])
def open_excel():
    if API_KEY and request.headers.get('x-bridge-key') != API_KEY:
        return jsonify({"success": False, "error": "Unauthorized"}), 401
    try:
        body = request.get_json(); data = body.get('data', {}); info = data.get('info', {}); eligibility = data.get('eligibility', {}); cert = data.get('certification', {})
        log_it(f"Exporting: {info.get('lname', 'Unknown')}")
        if not os.path.exists(TEMPLATE_PATH): return jsonify({"success": False, "error": "Template missing"}), 500
        wb = openpyxl.load_workbook(TEMPLATE_PATH); sn = [s.upper() for s in wb.sheetnames]
        ws_front = wb[wb.sheetnames[sn.index('FRONT')]] if 'FRONT' in sn else wb.active
        ws_back = wb[wb.sheetnames[sn.index('BACK')]] if 'BACK' in sn else (wb[wb.sheetnames[1]] if len(wb.sheetnames)>1 else wb.active)

        # --- STUDENT INFO (FRONT) ---
        safe_write(ws_front, 'F8', info.get('lname', '')); safe_write(ws_front, 'Y8', info.get('fname', '')); safe_write(ws_front, 'AZ8', info.get('mname', ''))
        safe_write(ws_front, 'C9', info.get('lrn', '')); safe_write(ws_front, 'AN9', info.get('sex', '')); safe_write(ws_front, 'AA9', info.get('birthdate', ''))
        safe_write(ws_front, 'BH9', info.get('admissionDate', ''))
        
        # --- ELIGIBILITY (FRONT) ---
        safe_write(ws_front, 'Z14', eligibility.get('schoolName', '')); safe_write(ws_front, 'AW14', eligibility.get('schoolAddress', ''))
        if eligibility.get('hsCompleter'): safe_write(ws_front, 'A13', 'X')
        safe_write(ws_front, 'N13', eligibility.get('hsGenAve', ''))
        if eligibility.get('jhsCompleter'): safe_write(ws_front, 'S13', 'X')
        safe_write(ws_front, 'AH13', eligibility.get('jhsGenAve', '')); safe_write(ws_front, 'P14', eligibility.get('gradDate', ''))
        if eligibility.get('pept'): safe_write(ws_front, 'A16', 'X')
        safe_write(ws_front, 'K16', eligibility.get('peptRating', ''))
        if eligibility.get('als'): safe_write(ws_front, 'S16', 'X')
        safe_write(ws_front, 'AC16', eligibility.get('alsRating', ''))
        if eligibility.get('others'): safe_write(ws_front, 'AH16', 'X')
        safe_write(ws_front, 'AP16', eligibility.get('othersSpec', '')); safe_write(ws_front, 'P17', eligibility.get('examDate', ''))
        safe_write(ws_front, 'AN17', eligibility.get('clcName', ''))

        # ==========================================================
        # GRADE 11 (FRONT PAGE)
        # ==========================================================
        
        # --- SEMESTER 1 (G11 - SEM 1) ---
        s1 = data.get('semester1', {})
        safe_write(ws_front, 'E23', s1.get('school', '')); safe_write(ws_front, 'AF23', s1.get('schoolId', ''))
        safe_write(ws_front, 'AS23', s1.get('gradeLevel', '')); safe_write(ws_front, 'BA23', s1.get('sy', ''))
        safe_write(ws_front, 'BK23', s1.get('semester', '')); safe_write(ws_front, 'G24', s1.get('trackStrand', ''))
        safe_write(ws_front, 'BC24', s1.get('section', ''))

        sub1 = s1.get('subjects', [])
        for i, r in enumerate(range(31, 42)): # 15 rows
            d = sub1[i] if i < len(sub1) else {}
            if d:
                safe_write(ws_front, f'A{r}', d.get('type','')); safe_write(ws_front, f'I{r}', d.get('subject',''))
                safe_write(ws_front, f'AT{r}', d.get('q1','')); safe_write(ws_front, f'AY{r}', d.get('q2',''))
                safe_write(ws_front, f'BD{r}', d.get('final','')); safe_write(ws_front, f'BI{r}', d.get('action',''))

        safe_write(ws_front, 'BD46', s1.get('genAve', ''))
        safe_write(ws_front, 'A47', s1.get('remarks', ''))

        r1 = s1.get('remedial', {})
        if r1:
            safe_write(ws_front, 'K50', r1.get('from', '')); safe_write(ws_front, 'U50', r1.get('to', ''))
            safe_write(ws_front, 'AF50', r1.get('school', '')); safe_write(ws_front, 'BA50', r1.get('schoolId', ''))
            for i, r in enumerate(range(52, 57)): # 5 remedial rows
                rd = r1.get('subjects', [])[i] if i < len(r1.get('subjects', [])) else {}
                if rd:
                    safe_write(ws_front, f'I{r}', rd.get('subject',''))
                    safe_write(ws_front, f'AT{r}', rd.get('semGrade','')); safe_write(ws_front, f'AY{r}', rd.get('remedialMark',''))
                    safe_write(ws_front, f'BD{r}', rd.get('recomputedGrade','')); safe_write(ws_front, f'BI{r}', rd.get('action',''))

        # --- SEMESTER 2 (G11 - SEM 2) ---
        s2 = data.get('semester2', {})
        safe_write(ws_front, 'E60', s2.get('school', '')); safe_write(ws_front, 'AF60', s2.get('schoolId', ''))
        safe_write(ws_front, 'AS60', s2.get('gradeLevel', '')); safe_write(ws_front, 'BA60', s2.get('sy', ''))
        safe_write(ws_front, 'BK60', s2.get('semester', '')); safe_write(ws_front, 'G61', s2.get('trackStrand', ''))
        safe_write(ws_front, 'BC61', s2.get('section', ''))

        sub2 = s2.get('subjects', [])
        for i, r in enumerate(range(68, 83)): # 15 rows
            d = sub2[i] if i < len(sub2) else {}
            if d:
                safe_write(ws_front, f'A{r}', d.get('type','')); safe_write(ws_front, f'I{r}', d.get('subject',''))
                safe_write(ws_front, f'AT{r}', d.get('q1','')); safe_write(ws_front, f'AY{r}', d.get('q2',''))
                safe_write(ws_front, f'BD{r}', d.get('final','')); safe_write(ws_front, f'BI{r}', d.get('action',''))

        safe_write(ws_front, 'BD83', s2.get('genAve', ''))
        safe_write(ws_front, 'A84', s2.get('remarks', ''))

        r2 = s2.get('remedial', {})
        if r2:
            safe_write(ws_front, 'K87', r2.get('from', '')); safe_write(ws_front, 'U87', r2.get('to', ''))
            safe_write(ws_front, 'AF87', r2.get('school', '')); safe_write(ws_front, 'BA87', r2.get('schoolId', ''))
            for i, r in enumerate(range(89, 94)): 
                rd = r2.get('subjects', [])[i] if i < len(r2.get('subjects', [])) else {}
                if rd:
                    safe_write(ws_front, f'I{r}', rd.get('subject',''))
                    safe_write(ws_front, f'AT{r}', rd.get('semGrade','')); safe_write(ws_front, f'AY{r}', rd.get('remedialMark',''))
                    safe_write(ws_front, f'BD{r}', rd.get('recomputedGrade','')); safe_write(ws_front, f'BI{r}', rd.get('action',''))

        # ==========================================================
        # GRADE 12 (BACK PAGE)
        # ==========================================================
        
        # --- SEMESTER 3 (G12 - SEM 1) ---
        s3 = data.get('semester3', {})
        safe_write(ws_back, 'E23', s3.get('school', '')); safe_write(ws_back, 'AF23', s3.get('schoolId', ''))
        safe_write(ws_back, 'AS23', s3.get('gradeLevel', '')); safe_write(ws_back, 'BA23', s3.get('sy', ''))
        safe_write(ws_back, 'BK23', s3.get('semester', '')); safe_write(ws_back, 'G24', s3.get('trackStrand', ''))
        safe_write(ws_back, 'BC24', s3.get('section', ''))

        sub3 = s3.get('subjects', [])
        for i, r in enumerate(range(31, 46)):
            d = sub3[i] if i < len(sub3) else {}
            if d:
                safe_write(ws_back, f'A{r}', d.get('type','')); safe_write(ws_back, f'I{r}', d.get('subject',''))
                safe_write(ws_back, f'AT{r}', d.get('q1','')); safe_write(ws_back, f'AY{r}', d.get('q2',''))
                safe_write(ws_back, f'BD{r}', d.get('final','')); safe_write(ws_back, f'BI{r}', d.get('action',''))

        safe_write(ws_back, 'BD46', s3.get('genAve', ''))
        safe_write(ws_back, 'A47', s3.get('remarks', ''))

        r3 = s3.get('remedial', {})
        if r3:
            safe_write(ws_back, 'K50', r3.get('from', '')); safe_write(ws_back, 'U50', r3.get('to', ''))
            safe_write(ws_back, 'AF50', r3.get('school', '')); safe_write(ws_back, 'BA50', r3.get('schoolId', ''))
            for i, r in enumerate(range(52, 57)):
                rd = r3.get('subjects', [])[i] if i < len(r3.get('subjects', [])) else {}
                if rd:
                    safe_write(ws_back, f'I{r}', rd.get('subject',''))
                    safe_write(ws_back, f'AT{r}', rd.get('semGrade','')); safe_write(ws_back, f'AY{r}', rd.get('remedialMark',''))
                    safe_write(ws_back, f'BD{r}', rd.get('recomputedGrade','')); safe_write(ws_back, f'BI{r}', rd.get('action',''))

        # --- SEMESTER 4 (G12 - SEM 2) ---
        s4 = data.get('semester4', {})
        safe_write(ws_back, 'E60', s4.get('school', '')); safe_write(ws_back, 'AF60', s4.get('schoolId', ''))
        safe_write(ws_back, 'AS60', s4.get('gradeLevel', '')); safe_write(ws_back, 'BA60', s4.get('sy', ''))
        safe_write(ws_back, 'BK60', s4.get('semester', '')); safe_write(ws_back, 'G61', s4.get('trackStrand', ''))
        safe_write(ws_back, 'BC61', s4.get('section', ''))

        sub4 = s4.get('subjects', [])
        for i, r in enumerate(range(68, 83)):
            d = sub4[i] if i < len(sub4) else {}
            if d:
                safe_write(ws_back, f'A{r}', d.get('type','')); safe_write(ws_back, f'I{r}', d.get('subject',''))
                safe_write(ws_back, f'AT{r}', d.get('q1','')); safe_write(ws_back, f'AY{r}', d.get('q2',''))
                safe_write(ws_back, f'BD{r}', d.get('final','')); safe_write(ws_back, f'BI{r}', d.get('action',''))

        safe_write(ws_back, 'BD83', s4.get('genAve', ''))
        safe_write(ws_back, 'A84', s4.get('remarks', ''))

        r4 = s4.get('remedial', {})
        if r4:
            safe_write(ws_back, 'K87', r4.get('from', '')); safe_write(ws_back, 'U87', r4.get('to', ''))
            safe_write(ws_back, 'AF87', r4.get('school', '')); safe_write(ws_back, 'BA87', r4.get('schoolId', ''))
            for i, r in enumerate(range(89, 94)):
                rd = r4.get('subjects', [])[i] if i < len(r4.get('subjects', [])) else {}
                if rd:
                    safe_write(ws_back, f'I{r}', rd.get('subject',''))
                    safe_write(ws_back, f'AT{r}', rd.get('semGrade','')); safe_write(ws_back, f'AY{r}', rd.get('remedialMark',''))
                    safe_write(ws_back, f'BD{r}', rd.get('recomputedGrade','')); safe_write(ws_back, f'BI{r}', rd.get('action',''))

        # --- CERTIFICATION ---
        safe_write(ws_back, 'G270', cert.get('gradDate', '')); safe_write(ws_back, 'G274', cert.get('genAve', '')); safe_write(ws_back, 'L270', cert.get('schoolHead', ''))
        
        filename = f"SF10_{str(info.get('lname','Student')).replace(' ','_')}_{int(time.time())}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, filename); wb.save(output_path)
        if sys.platform == 'win32': os.startfile(output_path)
        return jsonify({"success": True, "filePath": output_path})
    except Exception as e:
        log_it(f"ERR: {str(e)}"); return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=PORT, debug=False)
