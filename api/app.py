from flask import Flask, request, jsonify, render_template
from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from werkzeug.utils import secure_filename
from datetime import datetime
import uuid
import threading
import pathlib
import logging
import os
from explainer.GptExplainer import MainFunc

app = Flask(__name__, template_folder='../templates')
app.config['SECRET_KEY'] = 'supersecretkey'
app.config['UPLOAD_DIR'] = pathlib.Path(os.getcwd()) / 'uploads'
app.config['RESULT_DIR'] = pathlib.Path(os.getcwd()) / 'results'

def create_unique_filename(original_name: str, extension: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    unique_id = str(uuid.uuid4())[:8]
    return f"{original_name}_{timestamp}_{unique_id}{extension}"

class FileUploadForm(FlaskForm):
    document = FileField('Document')
    upload = SubmitField("Submit")

def process_file_async(file_path):
    result = MainFunc(file_path)
    print(f"Processing {'succeeded' if result == 1 else 'failed'} for: {file_path}")

@app.route('/')
def home():
    return render_template('welcome.html')

@app.route('/upload', methods=['GET', 'POST'])
def handle_upload():
    form = FileUploadForm()
    if request.method == 'POST' and form.validate_on_submit():
        uploaded_file = form.document.data
        original_filename = secure_filename(uploaded_file.filename)
        logging.info(f"Received file: {original_filename}")
        
        if not original_filename.lower().endswith(".pptx"):
            return jsonify({"error": "Unsupported file format. Please upload a PPTX file."}), 400
        
        unique_filename = create_unique_filename(*os.path.splitext(original_filename))
        file_path = app.config['UPLOAD_DIR'] / unique_filename
        uploaded_file.save(file_path)
        
        threading.Thread(target=process_file_async, args=(file_path,)).start()
        return jsonify({"id": unique_filename.split('_')[-1].split('.')[0]})
    
    return render_template('index.html', form=form)

@app.route('/status')
def get_status():
    file_id = request.args.get('id')
    if not file_id:
        return render_template('status_result.html', error="Missing file ID")
    
    for file in app.config['UPLOAD_DIR'].glob(f"*_{file_id}*"):
        original_name = "_".join(file.name.split("_")[:-2])
        upload_time = file.name.split("_")[-2]
        result_file = app.config['RESULT_DIR'] / f"{file.stem}.json"
        
        if result_file.exists():
            with open(result_file) as f:
                content = f.read()
            return render_template('status_result.html', status="complete", name=original_name, time=upload_time, content=content)
        else:
            return render_template('status_result.html', status="processing", name=original_name, time=upload_time)
    
    return render_template('status_result.html', error="File not found")

if __name__ == '__main__':
    app.config['UPLOAD_DIR'].mkdir(exist_ok=True)
    app.config['RESULT_DIR'].mkdir(exist_ok=True)
    app.run(port=5000)