from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
from pdf2docx import Converter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    converted_file_path = convert_pdf_to_word(file_path)
    
    return send_file(converted_file_path, as_attachment=True)

def convert_pdf_to_word(pdf_path):
    word_path = pdf_path.replace('.pdf', '.docx')
    cv = Converter(pdf_path)
    cv.convert(word_path, start=0, end=None)
    cv.close()
    return word_path

if __name__ == '__main__':
    app.run(debug=True)
