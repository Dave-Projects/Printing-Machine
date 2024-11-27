from flask import Flask, render_template, request, redirect, url_for
import os

# Web app
app = Flask(__name__)

# Specify the custom upload folder path
app.config['UPLOAD_FOLDER'] = 'C:\PrintVendoFiles\WifiReceived'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file:
        # Use the custom upload folder path
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)
        return render_template('upload.html', filename=filename)

# Handle non-existing routes and /upload by redirecting to /
@app.errorhandler(404)
@app.errorhandler(500)
@app.route('/upload')
def page_not_found(e):
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=80)
