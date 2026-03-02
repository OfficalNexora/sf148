import os
import sys
import json
import time
import datetime
import subprocess
from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl

# --- CONFIGURATION ---
PORT = 8787
API_KEY = "sf10-bridge-2024"
TEMPLATE_NAME = 'PLACEHOLDER(ALL).xlsx'

# Robust path detection for PyInstaller
if hasattr(sys, '_MEIPASS'):
    # Running from PyInstaller bundle
    EXE_DIR = os.path.dirname(sys.executable)
else:
    # Running from source
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

# Logging to help debug crashes
LOG_FILE = os.path.join(EXE_DIR, "excel-bridge-python.log")
def log_it(msg):
    try:
        with open(LOG_FILE, "a") as f:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{ts}] {msg}\n")
    except:
        pass

log_it(f"Bridge init. Template: {TEMPLATE_PATH}")

OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "form137-exports")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

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
        return jsonify({"success": False, "error": "Unauthorized"}), 401

    try:
        body = request.get_json()
        data = body.get('data', {})
        info = data.get('info', {})
        eligibility = data.get('eligibility', {})
        cert = data.get('certification', {})
        
        if not os.path.exists(TEMPLATE_PATH):
            log_it(f"ERR: Template not found at {TEMPLATE_PATH}")
            return jsonify({"success": False, "error": f"Template missing: {TEMPLATE_PATH}"}), 500

        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        
        # Robust sheet selection
        sn = [s.upper() for s in wb.sheetnames]
        ws_front = wb[wb.sheetnames[sn.index('FRONT')]] if 'FRONT' in sn else wb.active
        ws_back = wb[wb.sheetnames[sn.index('BACK')]] if 'BACK' in sn else (wb[wb.sheetnames[1]] if len(wb.sheetnames)>1 else wb.active)

        # --- STUDENT INFO ---
        ws_front['W7'] = str(info.get('lname', ''))
        ws_front['AC7'] = str(info.get('fname', ''))
        ws_front['AI7'] = str(info.get('mname', ''))
        ws_front['W9'] = str(info.get('lrn', ''))
        ws_front['AH9'] = str(info.get('sex', ''))
        ws_front['W11'] = str(info.get('birthdate', ''))
        ws_front['AF11'] = str(info.get('admissionDate', ''))
        
        # --- ELIGIBILITY ---
        ws_front['H22'] = str(eligibility.get('schoolName', ''))
        ws_front['AQ22'] = str(eligibility.get('schoolAddress', ''))
        if eligibility.get('pept'): ws_front['E20'] = 'X'
        ws_front['AQ18'] = str(eligibility.get('jhsGenAve', ''))

        # --- SEMESTER 1 ---
        s1 = data.get('semester1', {})
        ws_front['E23'] = str(s1.get('school', ''))
        ws_front['AC23'] = str(s1.get('schoolId', ''))
        ws_front['AP23'] = str(s1.get('gradeLevel', ''))
        ws_front['BA23'] = str(s1.get('sy', ''))
        ws_front['BK23'] = str(s1.get('semester', ''))
        
        for i, subj in enumerate(s1.get('subjects', [])):
            r = 31 + i
            ws_front[f'A{r}'] = str(subj.get('type', ''))
            ws_front[f'E{r}'] = str(subj.get('subject', ''))
            ws_front[f'AT{r}'] = str(subj.get('q1', ''))
            ws_front[f'AY{r}'] = str(subj.get('q2', ''))
            ws_front[f'BD{r}'] = str(subj.get('final', ''))
            ws_front[f'BI{r}'] = str(subj.get('actionTaken', ''))

        # --- BACK PAGE ---
        ws_back['G270'] = str(cert.get('gradDate', ''))
        ws_back['G274'] = str(cert.get('genAve', ''))
        ws_back['L270'] = str(cert.get('schoolHead', ''))
        
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
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    log_it(f"Starting server on port {PORT}...")
    app.run(host='127.0.0.1', port=PORT, debug=False)
