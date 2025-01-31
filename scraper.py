from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import logging
import time
import re
import os
from flask import Flask

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "C:/Users/Madiyar/Desktop/GQ/scraper_final/uploads"

app.config['OUTPUT_FOLDER'] = 'C:/Users/Madiyar/Desktop/GQ/scraper_final/outputs'



logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920x1080")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36")
chrome_options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def clean_price(price_text):
    match = re.search(r'(\d[\d\s]*[. ,]?\d{0,2})(₸|тг|KZT)', price_text)
    if match:
        number = match.group(1).replace(' ', '')
        currency = match.group(2)
        return f"{number} {currency}"
    return "Цена по запросу"
def scraper(target_url, query):
    search_url = f"{target_url}{query}"
    driver.get(search_url)
    time.sleep(3) 

    product_prices = []

    try:
        if 'nur-electro.kz' in target_url:
            product_selector = (By.CLASS_NAME, 'products')
            price_selector = (By.CLASS_NAME, 'price')
        elif '220volt.kz' in target_url:
            product_selector = (By.CLASS_NAME, 'cards__list')
            price_selector = (By.CLASS_NAME, 'product__buy-info-price-actual_value')
        elif 'ekt.kz' in target_url:
            product_selector = (By.CLASS_NAME, 'row')
            price_selector = (By.CSS_SELECTOR, '.left-block .price')
        elif 'barlau.kz' in target_url:
            product_selector = (By.CLASS_NAME, 'catalog-section-items')
            price_selector = (By.XPATH, "//span[@data-role='item.price.discount']")
        elif 'intant.kz' in target_url:
            product_selector = (By.CLASS_NAME, 'product_card__block_item_inner')
            price_selector = (By.CLASS_NAME, 'product-card-inner__new-price')
        elif 'elcentre.kz' in target_url:
            product_selector = (By.CLASS_NAME, 'b-product-gallery')
            price_selector = (By.CLASS_NAME, 'b-product-gallery__current-price')
        else:
            raise ValueError("URL неподдерживается")

        WebDriverWait(driver, 15).until(EC.presence_of_element_located(product_selector))
        products = driver.find_elements(*product_selector)

        for product in products:
            try:
                price_text = product.find_element(*price_selector).text
                logging.info(f"Цена: {price_text}")
                
                cleaned_price = clean_price(price_text)
                product_prices.append(cleaned_price)
            except NoSuchElementException:
                print 
            except Exception as e:
                logging.error(f"Error processing price: {str(e)}")
                product_prices.append("Error processing price")

    except Exception as e:
        logging.error(f"Не найдено")

    logging.info(f"Цены после парсинга: {product_prices}")
    return product_prices

def merge_excel_files(parsing_file, scraped_data, output_file, target_urls):
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        dfs = pd.read_excel(parsing_file, sheet_name=None)
        
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            if sheet_name in scraped_data:
                scraped_df = pd.DataFrame(scraped_data[sheet_name], columns=["Артикул"] + [f"Сайт {url}" for url in target_urls])
                scraped_df.to_excel(writer, sheet_name=sheet_name, startcol=len(df.columns), index=False)
        
        logging.info(f"Мердж {output_file}")

def main():
    dfs = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], 'test.xlsx'), sheet_name=None)
    logging.info(f"Sheets: {dfs.keys()}")
    
    search_queries = {}
    for sheet, df in dfs.items():
        if 'Артикул' in df.columns:
            search_queries[sheet] = df['Артикул'].dropna().tolist()

    logging.info(f"Search queries: {search_queries}")

    target_urls = [
        "https://220volt.kz/search?query=",
        "https://ekt.kz/catalog/?q=",
        "https://barlau.kz/catalog/?q=",
        "https://elcentre.kz/site_search?search_term=",
        "https://intant.kz/catalog/?q="
    ]

    final_data = {sheet: [] for sheet in search_queries}

    for sheet, queries in search_queries.items():
        for query in queries:
            row = [query]
            for target_url in target_urls:
                logging.info(f"Парсинг: {target_url}{query}")
                try:
                    prices = scraper(target_url, query)
                    row.append(", ".join(prices) if prices else "Не найдено")
                except Exception as e:
                    logging.error(f"Failed to scrape {target_url}{query}: {e}")
                    row.append(f"Error: {e}")
            final_data[sheet].append(row)

    merge_excel_files(
        os.path.join(app.config['UPLOAD_FOLDER'], 'test.xlsx'),
        final_data,
        os.path.join(app.config['OUTPUT_FOLDER'], 'merged.xlsx'),
        target_urls
    )
    logging.info("Excel files merged successfully")

    logging.info(f"Final data: {final_data}")

if __name__ == "__main__":
    try:
        main()
    finally:
        driver.quit()

# Terminal:
# pip install xlsxwriter