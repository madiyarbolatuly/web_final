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
import re
import os
import time
from flask import Flask
from urllib.parse import urlparse

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
chrome_options.add_argument('--ignore-certificate-errors')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def clean_price(price_text):
    match = re.search(r'(\d[\d\s]*[.,]?\d{0,2})', price_text)
    if match:
        number = match.group(1).replace(' ', '').replace(',', '.')
        return f"{number} "
    return "Цена по запросу"

def get_selectors(target_url):
    selectors = {
        'nur-electro.kz': ((By.CLASS_NAME, 'products'), (By.CLASS_NAME, 'price')),
        '220volt.kz': ((By.CLASS_NAME, 'cards__list'), (By.CLASS_NAME, 'product__buy-info-price-actual_value')),
        'ekt.kz': ((By.CLASS_NAME, 'row'), (By.CSS_SELECTOR, '.left-block .price')),
        'barlau.kz': ((By.CLASS_NAME, 'catalog-section-items'), (By.XPATH, "//span[@data-role='item.price.discount']")),
        'intant.kz': ((By.CLASS_NAME, 'product_card__block_item_inner'), (By.CLASS_NAME, 'product-card-inner__new-price')),
        'euroelectric.kz': ((By.CLASS_NAME, 'product-item'), (By.CLASS_NAME, 'product-price')),
        'albion-group.kz': ((By.CLASS_NAME, 'cs-product-gallery'), (By.CSS_SELECTOR, "span.cs-goods-price__value.cs-goods-price__value_type_current")),
        'chipdip.kz': ((By.CLASS_NAME, 'price.price-main'), (By.CSS_SELECTOR, "span[id^='price_']")),
        'volt.kz': ((By.CLASS_NAME, 'multi-snippet'), (By.XPATH, "//div[@class='multi-snippet']/span[@class='multi-price']")),
        'legrand24.kz': ((By.CLASS_NAME, 'summary entry-summary"'), (By.CLASS_NAME, 'woocommerce-Price-amount amount"'))
    }
    for key in selectors:
        if key in target_url:
            return selectors[key]
    raise ValueError("URL неподдерживается")

def scrape_prices(target_url, query):
    search_url = f"{target_url}{query}"
    driver.get(search_url)
    time.sleep(2)

    product_prices = []
    try:
        product_selector, price_selector = get_selectors(target_url)
        time.sleep(2)
        products = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(product_selector)
        )

        for product in products:
            try:
                product_link = product.find_element(By.TAG_NAME, 'a').get_attribute('href')
                if product_link:
                    driver.get(product_link)
                    time.sleep(2)
                    price_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located(price_selector)
                    )
                    price_text = price_element.text
                    # Extract text from nested elements if present
                    for child in price_element.find_elements(By.XPATH, ".//*"):
                        price_text = price_text.replace(child.text, "")
                else:
                    price_element = product.find_element(*price_selector)
                    price_text = price_element.text
                    # Extract text from nested elements if present
                    for child in price_element.find_elements(By.XPATH, ".//*"):
                        price_text = price_text.replace(child.text, "")
                
                cleaned_price = clean_price(price_text)
                product_prices.append(cleaned_price)
            except NoSuchElementException:
                continue
            except Exception as e:
                logging.error(f"Error processing price")
                #product_prices.append("Не найдено")

    except TimeoutException:
        logging.error(f"Error scraping prices: {e}")

#logging.info(f"Цены после парсинга: {product_prices}")
    #logging.info(f"Цены после парсинга: {product_prices}")
    return product_prices

def merge_excel_files(parsing_file, scraped_data, output_file, target_urls):
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        dfs = pd.read_excel(parsing_file, sheet_name=None)
        
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            if sheet_name in scraped_data:
                domain_names = [urlparse(url).netloc for url in target_urls]
                scraped_df = pd.DataFrame(scraped_data[sheet_name], columns=['Артикул'] + domain_names)
                scraped_df.to_excel(writer, sheet_name=sheet_name, startcol=len(df.columns), index=False)
        
        #logging.info(f"Merge {output_file}")

def main():
    dfs = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], 'test.xlsx'), sheet_name=None)
    #logging.info(f"Sheets: {dfs.keys()}")
    
    search_queries = {sheet: df['Артикул'].dropna().tolist() for sheet, df in dfs.items() if 'Артикул' in df.columns}
    #logging.info(f"Search queries: {search_queries}")

    final_data = {sheet: [] for sheet in search_queries}

    for sheet, queries in search_queries.items():
        for query in queries:
            row = [query]
            for target_url in target_urls:
                #logging.info(f"Сайт URL: {target_url}{query}")
                try:
                    prices = scrape_prices(target_url, query)
                    logging.info(f"{prices}")

                    row.append(", ".join(prices) if prices else "Не найдено")
                except Exception as e:
                    logging.error(f"Failed to scrape {target_url}{query}: {e}")
                    
            final_data[sheet].append(row)

    merge_excel_files(
        os.path.join(app.config['UPLOAD_FOLDER'], 'test.xlsx'),
        final_data,
        os.path.join(app.config['OUTPUT_FOLDER'], 'merged.xlsx'),
        target_urls
    )
    #logging.info("Excel files merged successfully")

    #logging.info(f"Final data: {final_data}")

if __name__ == "__main__":
    try:
        main()
    finally:
        driver.quit()

# Export target_urls for use in other files
target_urls = [
    #"https://220volt.kz/search?query=",
    #"https://ekt.kz/catalog/?q=",
    #"https://barlau.kz/catalog/?q=",
    #"https://elcentre.kz/site_search?search_term=",
    #"https://intant.kz/catalog/?q=",
    "https://albion-group.kz/site_search?search_term=",
    "https://www.chipdip.kz/search?searchtext=",
    """"https://volt.kz/#/search/"  can not access through the search_query""",
]