from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from download_process import download_and_process_video

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif'}
app.static_folder = 'static'



@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    # Get form data
    youtube_link = request.form['youtube_link']
    folder_name = request.form['folder_name']
    ppt_name = request.form['ppt_name']

    # Create a unique folder path for each user
    user_folder_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(folder_name))
    os.makedirs(user_folder_path, exist_ok=True)

    # Download and process the video
    download_and_process_video(youtube_link, user_folder_path, ppt_name)

    # Zip the folder
    zip_filename = f'{folder_name}.zip'
    zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
    os.system(f"zip -r {zip_path} {user_folder_path}")

    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
