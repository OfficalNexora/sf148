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

# Detect if we are running as a compiled EXE or a script
if getattr(sys, 'frozen', False):
    # If running as EXE, look in the same folder as the EXE
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # If running as script, look in the scripts folder or the parent release folder
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Priority 1: Next to the EXE/Script
# Priority 2: In a 'release' subfolder
# Priority 3: Fallback to name only
POSSIBLE_PATHS = [
    os.path.join(BASE_DIR, TEMPLATE_NAME),
    os.path.join(BASE_DIR, 'release', TEMPLATE_NAME),
    os.path.join(os.getcwd(), 'release', TEMPLATE_NAME),
    TEMPLATE_NAME
]

TEMPLATE_PATH = next((p for p in POSSIBLE_PATHS if os.path.exists(p)), POSSIBLE_PATHS[0])

# Logging setup (to help debug why the EXE fails)
LOG_FILE = os.path.join(BASE_DIR, 'excel-bridge-python.log')
def log_it(msg):
    try:
        with open(LOG_FILE, 'a') as f:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {msg}\n")
    except:
        pass

log_it(f"Bridge starting. Template path: {TEMPLATE_PATH}")

OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "form137-exports")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Flask App Setup ---
app = Flask(__name__)
CORS(app)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"success": True, "status": "Simplified Python Bridge Online", "template": TEMPLATE_PATH})

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
        
        log_it(f"Generating Excel for {info.get('lname', 'Unknown')}")

        if not os.path.exists(TEMPLATE_PATH):
            log_it(f"ERROR: Template missing at {TEMPLATE_PATH}")
            return jsonify({"success": False, "error": f"Template not found at {TEMPLATE_PATH}"}), 500
            
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        sheet_front = wb['FRONT'] if 'FRONT' in wb.sheetnames else wb.active
        sheet_back = wb['BACK'] if 'BACK' in wb.sheetnames else wb.active

        # --- HARDCODED MAPPINGS (FRONT) ---
        sheet_front['W7'] = str(info.get('lname', ''))
        sheet_front['AC7'] = str(info.get('fname', ''))
        sheet_front['AI7'] = str(info.get('mname', ''))
        sheet_front['W9'] = str(info.get('lrn', ''))
        sheet_front['AH9'] = str(info.get('sex', ''))
        sheet_front['W11'] = str(info.get('birthdate', ''))
        sheet_front['AF11'] = str(info.get('admissionDate', ''))
        
        sheet_front['H22'] = str(eligibility.get('schoolName', ''))
        sheet_front['AQ22'] = str(eligibility.get('schoolAddress', ''))
        if eligibility.get('pept'): sheet_front['E20'] = 'X'
        sheet_front['AQ18'] = str(eligibility.get('jhsGenAve', ''))

        # --- SEMESTER 1 (FRONT) ---
        s1 = data.get('semester1', {})
        sheet_front['E23'] = str(s1.get('school', ''))
        sheet_front['AC23'] = str(s1.get('schoolId', ''))
        sheet_front['AP23'] = str(s1.get('gradeLevel', ''))
        sheet_front['BA23'] = str(s1.get('sy', ''))
        sheet_front['BK23'] = str(s1.get('semester', ''))
        
        s1_subjects = s1.get('subjects', [])
        for i, subj in enumerate(s1_subjects):
            row = 31 + i
            sheet_front[f'A{row}'] = str(subj.get('type', ''))
            sheet_front[f'E{row}'] = str(subj.get('subject', ''))
            sheet_front[f'AT{row}'] = str(subj.get('q1', ''))
            sheet_front[f'AY{row}'] = str(subj.get('q2', ''))
            sheet_front[f'BD{row}'] = str(subj.get('final', ''))
            sheet_front[f'BI{row}'] = str(subj.get('actionTaken', ''))

        # --- BACK PAGE: CERTIFICATION ---
        sheet_back['G270'] = str(cert.get('gradDate', ''))
        sheet_back['G274'] = str(cert.get('genAve', ''))
        sheet_back['L270'] = str(cert.get('schoolHead', ''))
        
        # Save output
        last_name = str(info.get('lname', 'Student')).replace(' ', '_')
        timestamp = int(time.time())
        student_filename = f"SF10_{last_name}_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, student_filename)
        
        wb.save(output_path)
        log_it(f"Successfully saved to {output_path}")

        # Open automatically
        os.startfile(output_path)
        return jsonify({"success": True, "filePath": output_path})

    except Exception as e:
        log_it(f"CRITICAL ERROR: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    log_it(f"Server starting on http://127.0.0.1:{PORT}")
    app.run(host='127.0.0.1', port=PORT, debug=False)
