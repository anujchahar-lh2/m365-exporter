#!/usr/bin/env python3
"""
Flask server for M365 Exporter UI
Handles credential input, extraction requests, and progress tracking
"""

import json
import os
import subprocess
import threading
import time
from flask import Flask, request, jsonify
from pathlib import Path
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Store extraction status
extraction_status = {
    'active': False,
    'progress': 0,
    'current_user': '',
    'status': 'Idle',
    'errors': [],
    'start_time': None
}

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'ok'})

@app.route('/api/extract', methods=['POST'])
def start_extraction():
    """Start extraction with credentials and parameters"""
    try:
        # Check if extraction is already running
        if extraction_status['active']:
            return jsonify({'error': 'Extraction already in progress'}), 409

        # Get form data
        tenant_id = request.form.get('tenant_id', '').strip()
        client_id = request.form.get('client_id', '').strip()
        client_secret = request.form.get('client_secret', '').strip()
        admin_email = request.form.get('admin_email', '').strip()
        employee_emails = json.loads(request.form.get('employee_emails', '[]'))
        since_days = request.form.get('since_days', '')
        modified_after = request.form.get('modified_after', '')

        # Validate inputs
        if not all([tenant_id, client_id, client_secret]):
            return jsonify({'error': 'Missing Azure credentials'}), 400

        if not admin_email or '@' not in admin_email:
            return jsonify({'error': 'Invalid admin email'}), 400

        if not employee_emails:
            return jsonify({'error': 'No employee emails provided'}), 400

        for email in employee_emails:
            if '@' not in email:
                return jsonify({'error': f'Invalid email: {email}'}), 400

        # Create temp .env file
        env_path = os.path.join(os.path.dirname(__file__), '.env.temp')
        with open(env_path, 'w') as f:
            f.write(f"TENANT_ID={tenant_id}\n")
            f.write(f"CLIENT_ID={client_id}\n")
            f.write(f"CLIENT_SECRET={client_secret}\n")

        # Start extraction in background thread
        thread = threading.Thread(
            target=run_extraction,
            args=(admin_email, employee_emails, since_days, modified_after, env_path)
        )
        thread.daemon = True
        thread.start()

        return jsonify({
            'status': 'started',
            'message': f'Extraction started for {len(employee_emails)} employee(s)'
        }), 202

    except Exception as e:
        return jsonify({'error': str(e)}), 500

def run_extraction(admin_email, employee_emails, since_days, modified_after, env_path):
    """Run extraction command in background"""
    try:
        extraction_status['active'] = True
        extraction_status['progress'] = 0
        extraction_status['start_time'] = datetime.now()
        extraction_status['errors'] = []

        # Build command
        cmd = [
            'python3', 'run_m365_export.py',
            '--admin', admin_email,
            '--export-only',
            '--local-only'
        ]

        # Add employee emails
        for email in employee_emails:
            cmd.extend(['--only', email])

        # Add date range if provided
        if since_days:
            cmd.extend(['--since-days', since_days])
        if modified_after:
            cmd.extend(['--modified-after', modified_after])

        # Prepare environment with temp .env
        env = os.environ.copy()
        env['DOTENV_PATH'] = env_path

        # Run extraction
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=os.path.dirname(__file__),
            env=env
        )

        # Monitor progress
        total_lines = 0
        for line in iter(process.stdout.readline, ''):
            if line:
                total_lines += 1
                parse_progress_line(line, employee_emails)

                # Update progress based on output
                progress = min(95, (total_lines / 1000) * 100)
                extraction_status['progress'] = progress

                print(f"[PROGRESS] {progress:.0f}% - {line.strip()}")

        # Wait for completion
        stdout, stderr = process.communicate()

        if process.returncode == 0:
            extraction_status['status'] = 'Complete'
            extraction_status['progress'] = 100
            elapsed = time.time() - extraction_status['start_time'].timestamp()
            print(f"[SUCCESS] Extraction completed in {elapsed:.1f}s")
        else:
            extraction_status['status'] = 'Failed'
            extraction_status['errors'].append(stderr if stderr else 'Unknown error')
            print(f"[ERROR] Extraction failed: {stderr}")

    except Exception as e:
        extraction_status['status'] = 'Error'
        extraction_status['errors'].append(str(e))
        print(f"[ERROR] {e}")
    finally:
        extraction_status['active'] = False
        # Clean up temp .env
        try:
            os.remove(env_path)
        except:
            pass

def parse_progress_line(line, employee_emails):
    """Parse progress from log line"""
    line_lower = line.lower()

    # Check which user is being processed
    for email in employee_emails:
        if email in line:
            extraction_status['current_user'] = email.split('@')[0]
            break

    # Update status based on keywords
    if 'email' in line_lower and 'export' in line_lower:
        extraction_status['status'] = 'Exporting emails...'
    elif 'onedrive' in line_lower or 'file' in line_lower:
        extraction_status['status'] = 'Downloading OneDrive files...'
    elif 'pre-flight' in line_lower:
        extraction_status['status'] = 'Verifying credentials...'

@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current extraction status"""
    elapsed = None
    if extraction_status['start_time']:
        elapsed = time.time() - extraction_status['start_time'].timestamp()

    return jsonify({
        'active': extraction_status['active'],
        'progress': extraction_status['progress'],
        'current_user': extraction_status['current_user'],
        'status': extraction_status['status'],
        'errors': extraction_status['errors'],
        'elapsed_seconds': elapsed
    })

@app.route('/api/cancel', methods=['POST'])
def cancel_extraction():
    """Cancel ongoing extraction"""
    if extraction_status['active']:
        extraction_status['active'] = False
        extraction_status['status'] = 'Cancelled'
        return jsonify({'status': 'cancelled'}), 200
    return jsonify({'error': 'No extraction in progress'}), 400

@app.route('/')
def index():
    """Serve the UI"""
    ui_path = os.path.join(os.path.dirname(__file__), 'm365-exporter-ui.html')
    with open(ui_path, 'r') as f:
        return f.read()

if __name__ == '__main__':
    print("🚀 M365 Exporter Server starting...")
    print("📊 Open http://localhost:5000 in your browser")
    app.run(debug=False, host='127.0.0.1', port=5000)
