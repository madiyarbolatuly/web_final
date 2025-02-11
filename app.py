# filepath: /C:/Users/Madiyar/Desktop/GQ/scraper_final/scraper_final/src/app.py
from flask import Flask, render_template, request, send_file, jsonify
from flask_socketio import SocketIO, emit
import os
from scraper import main, scrape_prices, target_urls
import logging

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
socketio = SocketIO(app)

def create_folder_if_not_exists(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

create_folder_if_not_exists(app.config['UPLOAD_FOLDER'])
create_folder_if_not_exists(app.config['OUTPUT_FOLDER'])

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

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

@app.route('/search')
def search():
    return render_template('search.html')

@socketio.on('search_artikul')
def handle_search_artikul(data):
    artikul = data['artikul']
    results = []

    for url in target_urls:
        prices = scrape_prices(url, artikul)
        results.append({
            'artikul': artikul,
            'url': url + artikul,
            'price': ", ".join(prices) if prices else "Не найдено"
        })
        logging.info(f"Final data: {prices}")
        logging.info(f"Scraping completed for URL: {url}{artikul}")

    emit('search_results', results)

if __name__ == '__main__':
    socketio.run(app, debug=True)