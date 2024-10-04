from models.zip_manager import ZipFolderManager
import shutil
import os
from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for

app = Flask(__name__)

# Folder to save the processed files
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function that processes the zip file using FolderProcessor
def process_zip(file_path):
    # Create FolderProcessor object to process the file
    processor = ZipFolderManager(file_path)
    processor.clean_all()
    # Return the path of the processed zip file
    return processor.processed_zip

# Home route with upload form
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle the file upload and processing
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file part'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'No selected file'})
    
    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        # Process the uploaded zip file using FolderProcessor
        processed_zip_path = process_zip(file_path)
        
        return jsonify({'success': True, 'filename': os.path.basename(processed_zip_path)})

# Route to download the processed file
@app.route('/download_file/<filename>')
def download_file(filename):

    file_path = os.path.join(filename)
    return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=False, use_reloader=False)
