import os
import shutil
import subprocess
from flask import Flask, render_template, request, jsonify, send_from_directory

app = Flask(__name__)

UPLOAD_FOLDER = 'input_jobdesc_lama'
OUTPUT_FOLDER = 'output_jobdesc_baru'
DEPT_FILE = 'DepartmentNameTitle.txt'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f'Gagal menghapus {file_path}: {e}')

@app.route('/')
def index():
    output_files = os.listdir(app.config['OUTPUT_FOLDER'])
    has_files = len(output_files) > 0
    return render_template('index.html', has_files=has_files)

@app.route('/upload', methods=['POST'])
def upload_file():
    clear_folder(app.config['UPLOAD_FOLDER'])
    clear_folder(app.config['OUTPUT_FOLDER'])
    
    files = request.files.getlist('files')
    saved_count = 0
    for file in files:
        if file.filename != '':
            # MURNI NAMA FILE ASLI TANPA DIUBAH JADI UNDERSCORE OLEH SECURE_FILENAME!
            filename = file.filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            saved_count += 1
            
    return jsonify({'status': 'success', 'message': f'{saved_count} file berhasil di-upload!'})

@app.route('/save_dept', methods=['POST'])
def save_dept():
    data = request.get_json()
    dept_name = data.get('department_name', '')
    # Menyimpan teks mentah ke DepartmentNameTitle.txt TANPA TANDA PETIK SAMA SEKALI
    with open(DEPT_FILE, 'w', encoding='utf-8') as f:
        f.write(dept_name)
    return jsonify({'status': 'success'})

@app.route('/run_script', methods=['POST'])
def run_script():
    try:
        subprocess.run(
            ['python', 'main.py'], 
            capture_output=True, 
            text=True, 
            encoding='utf-8', 
            check=True
        )
        
        output_files = os.listdir(app.config['OUTPUT_FOLDER'])
        return jsonify({'status': 'success', 'files': output_files})
    except subprocess.CalledProcessError as e:
        print("--- ERROR DI MAIN.PY ---")
        print(e.stderr)
        return jsonify({'status': 'error', 'message': e.stderr}), 500

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    clear_folder(UPLOAD_FOLDER)
    clear_folder(OUTPUT_FOLDER)
    app.run(debug=True, port=5000)