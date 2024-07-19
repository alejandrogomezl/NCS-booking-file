from flask import Flask, request, redirect, url_for, send_from_directory, render_template
import os
import shutil
import zipfile
from read_hotelgest import read_excel
from read_new import read_csv
from write_NCS import write_excel

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'out'
ZIP_FOLDER = 'zip'
ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

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
        return "No se encontró el archivo en la solicitud", 400
    file = request.files['file']
    if file.filename == '':
        return "El nombre del archivo está vacío", 400
    if not allowed_file(file.filename):
        return "Archivo no permitido", 400
    if file:
        filename = file.filename
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Procesa el archivo
        input_file_path = file_path
        print(f"Archivo subido: {input_file_path}")
        
        if filename.endswith('.xlsx'):
            excel_data = read_excel(input_file_path).read_excel()
        elif filename.endswith('.csv'):
            excel_data = read_csv(input_file_path).read_csv()
        else:
            return "Formato de archivo no soportado", 400
        
        write_excel(excel_data).write()
        
        # Redirige a la página de descarga
        return redirect(url_for('download_file'))
    else:
        return "Error desconocido", 400

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
    # Vaciar y recrear las carpetas
    def empty_and_create(folder):
        if os.path.exists(folder):
            shutil.rmtree(folder)
        os.makedirs(folder)

    empty_and_create(UPLOAD_FOLDER)
    empty_and_create(OUTPUT_FOLDER)
    empty_and_create(ZIP_FOLDER)

    app.run(debug=True)