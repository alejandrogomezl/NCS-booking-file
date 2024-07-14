from flask import Flask, request, redirect, url_for, send_from_directory, render_template
import os
import zipfile
from read_hotelgest import read_excel
from write_NCS import write_excel

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'out'
ZIP_FOLDER = 'zip'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['ZIP_FOLDER'] = ZIP_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def upload_file():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file_post():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = "in1.xlsx"
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        
        # Procesa el archivo
        input_file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        print(f"Archivo subido: {input_file_path}")
        excel_data = read_excel(input_file_path).read_excel()
        write_excel(excel_data).write()
        
        # Redirige a la página de descarga
        return redirect(url_for('download_file'))

@app.route('/download')
def download_file():
    files = os.listdir(app.config['OUTPUT_FOLDER'])
    print(f"Archivos en la carpeta de salida: {files}")
    return render_template('download.html', files=files)

@app.route('/download_zip')
def download_zip():
    zip_filename = os.path.join(app.config['ZIP_FOLDER'], 'output_files.zip')
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for root, _, files in os.walk(app.config['OUTPUT_FOLDER']):
            for file in files:
                zipf.write(os.path.join(root, file), file)
    return send_from_directory(app.config['ZIP_FOLDER'], 'output_files.zip')

@app.route('/download/<filename>')
def download(filename):
    try:
        return send_from_directory(app.config['OUTPUT_FOLDER'], filename)
    except FileNotFoundError:
        return "Archivo no encontrado", 404

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
    if not os.path.exists(ZIP_FOLDER):
        os.makedirs(ZIP_FOLDER)
    app.run(debug=True)
