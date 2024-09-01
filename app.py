from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
import os
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(url_for('index'))

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    # Debugging: Print file path and check if file exists
    print(f"File uploaded to: {filepath}")
    if os.path.isfile(filepath):
        print("File exists.")

    # Read the Excel file into a DataFrame
    try:
        df = pd.read_excel(filepath)
        print("DataFrame read successfully")
        print("DataFrame head:")
        print(df.head())  # Print the first few rows to check

        # Replace NaN values with empty strings
        df = df.fillna('')

        # Convert DataFrame to HTML table
        table = df.to_html(classes='data', header="true", index=False)
        return render_template('index.html', table=table, download_url=url_for('download', filename=file.filename))
    except Exception as e:
        print(f"Error processing file: {e}")
        flash(f"Error processing file: {e}")
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        flash('File not found')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
