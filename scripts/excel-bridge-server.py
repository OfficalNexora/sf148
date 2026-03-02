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
# The user wants to use 'PLACEHOLDER(ALL).xlsx' as the master template
TEMPLATE_NAME = 'PLACEHOLDER(ALL).xlsx'

# Path to the template (looking in release folder)
TEMPLATE_PATH = os.path.join(os.getcwd(), 'release', TEMPLATE_NAME)

OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "form137-exports")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Flask App Setup ---
app = Flask(__name__)
CORS(app)

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"success": True, "status": "Simplified Python Bridge Online"})

@app.route('/open-excel', methods=['POST'])
def open_excel():
    if API_KEY and request.headers.get('x-bridge-key') != API_KEY:
        return jsonify({"success": False, "error": "Unauthorized"}), 401

    try:
        # Load Data from Website
        body = request.get_json()
        data = body.get('data', {})
        
        # Shortcuts for you to use in hardcoding
        info = data.get('info', {})
        eligibility = data.get('eligibility', {})
        cert = data.get('certification', {})
        
        # 1. Load the existing Form 137 template
        # 'data_only=False' ensures formulas stay as formulas
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"success": False, "error": f"Template not found at {TEMPLATE_PATH}"}), 500
            
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        
        # Get your sheets
        sheet_front = wb['FRONT'] if 'FRONT' in wb.sheetnames else wb.active
        sheet_back = wb['BACK'] if 'BACK' in wb.sheetnames else wb.active

        # 2. UPDATE ONLY THE INPUT BOXES (HARDCODE HERE)
        # You can change these addresses (e.g., 'W7', 'AC7') whenever you want!
        
        # --- FRONT PAGE: BASIC INFO ---
        sheet_front['W7'] = str(info.get('lname', ''))
        sheet_front['AC7'] = str(info.get('fname', ''))
        sheet_front['AI7'] = str(info.get('mname', ''))
        sheet_front['W9'] = str(info.get('lrn', ''))
        sheet_front['AH9'] = str(info.get('sex', ''))
        sheet_front['W11'] = str(info.get('birthdate', ''))
        sheet_front['AF11'] = str(info.get('admissionDate', ''))
        
        # --- FRONT PAGE: ELIGIBILITY ---
        sheet_front['H22'] = str(eligibility.get('schoolName', ''))
        sheet_front['AQ22'] = str(eligibility.get('schoolAddress', ''))
        if eligibility.get('pept'): sheet_front['E20'] = 'X'
        sheet_front['AQ18'] = str(eligibility.get('jhsGenAve', ''))

        # --- FRONT PAGE: SEMESTER 1 ---
        s1 = data.get('semester1', {})
        sheet_front['E23'] = str(s1.get('school', ''))
        sheet_front['AC23'] = str(s1.get('schoolId', ''))
        sheet_front['AP23'] = str(s1.get('gradeLevel', ''))
        sheet_front['BA23'] = str(s1.get('sy', ''))
        sheet_front['BK23'] = str(s1.get('semester', ''))
        
        # Semester 1 Grades (Example Loop)
        s1_subjects = s1.get('subjects', [])
        for i, subj in enumerate(s1_subjects):
            row = 31 + i # Starts at Row 31
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
        
        # 3. Save as a NEW file
        last_name = str(info.get('lname', 'Student')).replace(' ', '_')
        timestamp = int(time.time())
        student_filename = f"SF10_{last_name}_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, student_filename)
        
        wb.save(output_path)
        print(f"Successfully generated {student_filename} with original formatting!")

        # 4. Open the file for the user
        os.startfile(output_path)

        return jsonify({"success": True, "filePath": output_path})

    except Exception as e:
        print(f"Error: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    print(f"================================================")
    print(f"       SF10 PYTHON BRIDGE (ULTRA-SIMPLE)       ")
    print(f"================================================")
    print(f"Address: http://127.0.0.1:{PORT}")
    print(f"Template Required: {TEMPLATE_PATH}")
    app.run(host='127.0.0.1', port=PORT, debug=False)
