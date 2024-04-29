from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'UFILES'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
FILE_LIST_FILE = 'uploaded_files.txt'  # Name of the text file to store uploaded file paths

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('pag.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'})
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Clear the contents of the text file
        open(FILE_LIST_FILE, 'w').close()
        
        # Write the full paths of all uploaded files to the text file
        with open(FILE_LIST_FILE, 'a') as file_list:
            for root, dirs, files in os.walk(app.config['UPLOAD_FOLDER']):
                for file in files:
                    file_list.write(os.path.join(root, file) + '\n')
        
        return render_template('uploadSuccess.html')
    else:
        return jsonify({'error': 'Invalid file type'})

if __name__ == '__main__':
    app.run(debug=True)