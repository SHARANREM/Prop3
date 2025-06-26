from flask import Flask, request, send_file, jsonify, render_template_string
import os
import subprocess
import platform
import time
from PyPDF2 import PdfMerger
import uuid
import csv
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
MERGED_FOLDER = 'merged'
LOG_FILE = 'conversion_log.csv'

for folder in [UPLOAD_FOLDER, CONVERTED_FOLDER, MERGED_FOLDER]:
    os.makedirs(folder, exist_ok=True)

# Ensure CSV has headers if not present
if not os.path.exists(LOG_FILE):
    with open(LOG_FILE, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Timestamp', 'Filename', 'Type', 'Size_MB', 'ConversionTime_sec'])

# Read CSV rows
def read_logs():
    with open(LOG_FILE, 'r') as f:
        reader = csv.DictReader(f)
        return list(reader)[::-1]  # reverse to show latest first

# LibreOffice-based PDF conversion
def convert_to_pdf(input_path, output_folder):
    libreoffice_cmd = 'libreoffice'
    if platform.system() == "Windows":
        libreoffice_cmd = r"C:\Program Files\LibreOffice\program\soffice.exe"

    try:
        subprocess.run([
            libreoffice_cmd,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_folder,
            input_path
        ], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print("Conversion error:", e)
        return False

@app.route('/', methods=['GET'])
def upload_form():
    logs = read_logs()
    return render_template_string("""
        <html>
        <head>
            <title>Convert & Merge Files</title>
            <style>
                body { font-family: Arial; margin: 30px; }
                table { border-collapse: collapse; width: 100%; margin-top: 30px; }
                th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
                th { background-color: #f2f2f2; }
            </style>
        </head>
        <body>
            <h2>Upload .docx, .pptx, .xlsx files to convert and merge</h2>
            <form method="POST" action="/convert-merge" enctype="multipart/form-data">
                <input type="file" name="files" multiple required>
                <button type="submit">Convert & Merge</button>
            </form>

            <h2>📄 Conversion History</h2>
            {% if logs %}
            <table>
                <tr>
                    <th>Timestamp</th>
                    <th>Filename</th>
                    <th>Type</th>
                    <th>Size (MB)</th>
                    <th>Conversion Time (sec)</th>
                </tr>
                {% for row in logs %}
                <tr>
                    <td>{{ row['Timestamp'] }}</td>
                    <td>{{ row['Filename'] }}</td>
                    <td>{{ row['Type'] }}</td>
                    <td>{{ row['Size_MB'] }}</td>
                    <td>{{ row['ConversionTime_sec'] }}</td>
                </tr>
                {% endfor %}
            </table>
            {% else %}
            <p>No conversion logs found yet.</p>
            {% endif %}
        </body>
        </html>
    """, logs=logs)

@app.route('/convert-merge', methods=['POST'])
def convert_and_merge():
    if 'files' not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist('files')
    pdf_paths = []

    for file in files:
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ['.docx', '.pptx', '.xlsx']:
            return jsonify({"error": f"{file.filename} has unsupported format"}), 400

        uid = str(uuid.uuid4())
        input_path = os.path.join(UPLOAD_FOLDER, f"{uid}_{file.filename}")
        file.save(input_path)

        size_mb = round(os.path.getsize(input_path) / (1024 * 1024), 2)
        file_type = ext.strip(".")

        before = set(os.listdir(CONVERTED_FOLDER))
        start = time.time()

        if convert_to_pdf(input_path, CONVERTED_FOLDER):
            time.sleep(0.5)
            after = set(os.listdir(CONVERTED_FOLDER))
            new_files = list(after - before)
            matched_file = next((f for f in new_files if f.endswith('.pdf')), None)

            if matched_file:
                converted_path = os.path.join(CONVERTED_FOLDER, matched_file)
                pdf_paths.append(converted_path)
                duration = round(time.time() - start, 2)

                with open(LOG_FILE, 'a', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        file.filename,
                        file_type,
                        size_mb,
                        duration
                    ])
            else:
                return jsonify({"error": f"Converted PDF for {file.filename} not found"}), 500
        else:
            return jsonify({"error": f"Failed to convert {file.filename}"}), 500

    # Merge PDFs
    merger = PdfMerger()
    for path in pdf_paths:
        merger.append(path)

    final_pdf = os.path.join(MERGED_FOLDER, f'merged_{uuid.uuid4().hex}.pdf')
    merger.write(final_pdf)
    merger.close()

    return send_file(final_pdf, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8000))  # Render sets this
    app.run(host='0.0.0.0', port=port, debug=True)

