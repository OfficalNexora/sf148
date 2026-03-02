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

# Check multiple places for the template
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

# Output directory for exported files
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "form137-exports")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

# Logging to help debug crashes
LOG_FILE = os.path.join(EXE_DIR, "excel-bridge-python.log")
def log_it(msg):
    try:
        with open(LOG_FILE, "a") as f:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{ts}] {msg}\n")
    except:
        pass

# --- HELPER: SAFE WRITE (Handles Merged Cells) ---
def safe_write(ws, row_or_coord, col_or_val, val=None):
    """
    Usage:
      safe_write(ws, 'W7', "Hello")
      safe_write(ws, 31, 1, "Hello")
    """
    if val is None:
        coord = str(row_or_coord).upper()
        write_val = col_or_val
    else:
        coord = f"{get_column_letter(col_or_val)}{row_or_coord}"
        write_val = val

    try:
        ws[coord] = write_val
    except AttributeError:
        # Happens if coord is a MergedCell (read-only)
        for merged_range in ws.merged_cells.ranges:
            if coord in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = write_val
                return
        log_it(f"Failed to write to {coord} even after merge check")

app = Flask(__name__)
CORS(app)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        "success": True, 
        "status": "Running",
        "template": TEMPLATE_PATH,
        "exe_dir": EXE_DIR
    })

@app.route('/open-excel', methods=['POST'])
def open_excel():
    if API_KEY and request.headers.get('x-bridge-key') != API_KEY:
        log_it("Unauthorized request attempt.")
        return jsonify({"success": False, "error": "Unauthorized"}), 401

    try:
        body = request.get_json()
        data = body.get('data', {})
        info = data.get('info', {})
        eligibility = data.get('eligibility', {})
        cert = data.get('certification', {})
        
        log_it(f"Processing export for {info.get('lname', 'Unknown')}")

        if not os.path.exists(TEMPLATE_PATH):
            log_it(f"ERR: Template not found at {TEMPLATE_PATH}")
            return jsonify({"success": False, "error": f"Template missing: {TEMPLATE_PATH}"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        
        # Robust sheet selection
        sn = [s.upper() for s in wb.sheetnames]
        ws_front = wb[wb.sheetnames[sn.index('FRONT')]] if 'FRONT' in sn else wb.active
        ws_back = wb[wb.sheetnames[sn.index('BACK')]] if 'BACK' in sn else (wb[wb.sheetnames[1]] if len(wb.sheetnames)>1 else wb.active)

        # ==========================================================
        # HARDCODE YOUR MAPPINGS BELOW
        # ==========================================================
        
        # --- STUDENT INFO (FRONT) ---
        safe_write(ws_front, 'W7', info.get('lname', ''))
        safe_write(ws_front, 'AC7', info.get('fname', ''))
        safe_write(ws_front, 'AI7', info.get('mname', ''))
        safe_write(ws_front, 'W9', info.get('lrn', ''))
        safe_write(ws_front, 'AH9', info.get('sex', ''))
        safe_write(ws_front, 'W11', info.get('birthdate', ''))
        safe_write(ws_front, 'AF11', info.get('admissionDate', ''))
        
        # --- ELIGIBILITY (FRONT) ---
        safe_write(ws_front, 'H22', eligibility.get('schoolName', ''))
        safe_write(ws_front, 'AQ22', eligibility.get('schoolAddress', ''))
        if eligibility.get('pept'): safe_write(ws_front, 'E20', 'X')
        safe_write(ws_front, 'AQ18', eligibility.get('jhsGenAve', ''))

        # --- SEMESTER 1 (FRONT) ---
        s1 = data.get('semester1', {})
        safe_write(ws_front, 'E23', s1.get('school', ''))
        safe_write(ws_front, 'AC23', s1.get('schoolId', ''))
        safe_write(ws_front, 'AP23', s1.get('gradeLevel', ''))
        safe_write(ws_front, 'BA23', s1.get('sy', ''))
        safe_write(ws_front, 'BK23', s1.get('semester', ''))
        
        for i, subj in enumerate(s1.get('subjects', [])):
            row = 31 + i
            safe_write(ws_front, row, 1, subj.get('type', ''))
            safe_write(ws_front, row, 5, subj.get('subject', ''))
            safe_write(ws_front, row, 46, subj.get('q1', ''))
            safe_write(ws_front, row, 51, subj.get('q2', ''))
            safe_write(ws_front, row, 56, subj.get('final', ''))
            safe_write(ws_front, row, 61, subj.get('actionTaken', ''))

        # --- BACK PAGE (CERTIFICATION) ---
        safe_write(ws_back, 'G270', cert.get('gradDate', ''))
        safe_write(ws_back, 'G274', cert.get('genAve', ''))
        safe_write(ws_back, 'L270', cert.get('schoolHead', ''))
        
        # Save output
        last_name = str(info.get('lname', 'Student')).replace(' ', '_')
        ts = int(time.time())
        filename = f"SF10_{last_name}_{ts}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        
        wb.save(output_path)
        log_it(f"Generated: {filename}")

        if sys.platform == 'win32':
            os.startfile(output_path)
            
        return jsonify({"success": True, "filePath": output_path})

    except Exception as e:
        log_it(f"CRITICAL ERR: {str(e)}")
        import traceback
        log_it(traceback.format_exc())
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    log_it(f"Starting server on port {PORT}...")
    app.run(host='127.0.0.1', port=PORT, debug=False)
