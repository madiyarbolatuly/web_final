from flask import Flask, render_template, request, send_file
import os
from scraper import main
import time

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

if not os.path.exists(app.config['OUTPUT_FOLDER']):
    os.makedirs(app.config['OUTPUT_FOLDER'])

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['excel_file']
        if file and file.filename.endswith('.xlsx'):
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], 'test.xlsx')
            file.save(input_path)
            
        
            try:
                main()  
                output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'merged.xlsx')
                return send_file(output_path, as_attachment=True)
            except Exception as e:
                return f"Error: {str(e)}"
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)