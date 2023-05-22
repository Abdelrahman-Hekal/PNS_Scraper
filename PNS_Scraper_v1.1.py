from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver as uc
import time
import csv
import os
from datetime import datetime
import pandas as pd
import numpy as np
import unidecode
import warnings
import re
import sys
import math
import shutil
import xlsxwriter
warnings.filterwarnings('ignore')

def initialize_bot():
    
    print('Initializing the web driver ...')
    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument('--headless=new')
    chrome_options.page_load_strategy = 'normal'
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(10000)

    return driver

def process_links(driver, links, settings):

    print('-'*100)
    print('Processing links before scraping')
    print('-'*100)
    df = pd.DataFrame()
    prods_limit = settings['Product Limit']
    n = len(links)
    for i, link in enumerate(links):

        print(f'Processing input link {i+1}/{n}...')
        # single product link
        if '/p/' in link:
            df = df.append([{'Link':link, 'Input Link':link}])
            continue

        driver.get(link)
        time.sleep(3)
        nprods = 0   
        try:
            # handling lazy loading
            while True:  
                try:
                    height1 = driver.execute_script("return document.body.scrollHeight")
                    driver.execute_script(f"window.scrollTo(0, {height1})")
                    time.sleep(8)
                    height2 = driver.execute_script("return document.body.scrollHeight")
                    if height1 == height2: 
                        break
                    if prods_limit > 0:
                        titles = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[class='productName']")))
                        if len(titles) > prods_limit:
                            break
                except Exception as err:
                    break
            
            titles = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[class='productName']")))
            for title in titles:
                try:
                    url = title.get_attribute('href')
                    if '/p/' in url:
                        df = df.append([{'Link': url, 'Input Link':link}])
                        nprods += 1
                    if nprods == prods_limit:
                        break
                except:
                    pass

        except Exception as err:
            print(f"The below error occurred while scraping the products urls under link {i+1} \n")
            print(err)
            print('-'*100)
            driver.quit()
            time.sleep(5)
            driver = initialize_bot()
            driver.get(link)
            time.sleep(3)
            
    # return products links
    df.drop_duplicates(inplace=True)
    prod_links = df['Link'].values.tolist()
    input_links = df['Input Link'].values.tolist()
    return prod_links, input_links


def scrape_prods(driver, prod_links, input_links, output1, output2, settings):

    keys = ["Product ID","Product URL",	"Product Title", "Brand", "Price (HKD)","Product Origin","Product Category","Product Description","Product Delivery","Product Rating","Image URL", "Promotion", "Availability", "Input Link"]

    print('-'*100)
    print('Scraping links')
    print('-'*100)

    stamp = datetime.now().strftime("%d_%m_%Y")
    # reading scraped links for skipping
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['Product URL'].values.tolist()
    except:
        pass

    prods = pd.DataFrame()
    comments = pd.DataFrame()
    nlinks = len(prod_links)  
    for i, link in enumerate(prod_links):

        if link in scraped: continue
        prod = {}
        for key in keys:
            prod[key] = ''

        print(f'Scraping the details from link {i+1}/{nlinks} ...')
        driver.get(link)
        time.sleep(1)

        # handling 404 error
        try:
            wait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//meta[@content='404 Not Found | PARKnSHOP eShop']")))  
            print(f'Warning - 404 Error in link: {link}')
            continue
        except:
            pass
           
        # scrolling across the page 
        try:
            htmlelement= wait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
            total_height = driver.execute_script("return document.body.scrollHeight")
            height = total_height/30
            new_height = 0
            for _ in range(40):
                prev_hight = new_height
                new_height += height             
                driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                time.sleep(0.1)
        except:
            pass

        try:
            # scraping product URL
            prod['Product URL'] = link

            # scraping product ID
            ID = ''
            try:
                ID = link.split('/p/BP_')[-1]      
                prod['Product ID'] = ID   
            except:
                pass



            # scraping product title
            title, unit = '', ''
            try:
                title = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.product-name"))).get_attribute('textContent')
                prod['Product Title'] = title
                unit = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='product-unit']"))).get_attribute('textContent').strip()
                prod['Product Title'] = title + ' ' + f"({unit})"
            except:
                pass           
           
            # scraping product brand
            brand = ''
            try:
                brand = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='product-brand']"))).get_attribute('textContent')
                prod['Brand'] = brand
            except:
                pass
                
            # scraping Price (HKD)
            price = ''
            try:
                price_div = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.product-price-group")))
                try:
                    price = wait(price_div, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class*='currentPrice']"))).get_attribute('textContent').replace('$', '').replace(',', '').strip()
                except:
                    price = wait(price_div, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='currentPrice']"))).get_attribute('textContent').replace('$', '').replace(',', '').strip()

                prod['Price (HKD)'] = price
            except:
                pass

            # scraping product origion
            origin = ''
            try:
                origin_tile = wait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "pns-origin")))
                origin = wait(origin_tile, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.info-content"))).get_attribute('textContent')
                prod['Product Origin'] = origin.replace('<BR>', ' ').strip()       
            except:
                pass
            
            # scraping product delivery
            delivery = ''
            try:
                reg = wait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "pns-product-pickup")))
                methods = wait(reg, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.delivery-options")))  
                for method in methods:
                    try:
                        delivery += method.get_attribute('textContent') + '\n'
                    except:
                        pass
                    
                prod['Product Delivery'] = delivery.strip("\n")
            except:
                pass
                
            # scraping product description
            description = ''
            try:
                details = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.description-group")))  
                topic = wait(details, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.description-topic"))).get_attribute('textContent')
                description = description + topic.strip('"') + '\n'
                prod['Product Description'] = description.strip("\n") 
                divs = wait(details, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.detail")))
                for div in divs:
                    try:
                        title = wait(div, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h2.detail-title"))).get_attribute('textContent')
                        content = wait(div, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.detail-content"))).get_attribute('textContent')
                        description += title +' : '+ content + '\n'
                    except:
                        pass                             
            except:
                pass

            prod['Product Description'] = description.strip("\n") 

            # scraping product category
            cat = ''
            try:
                cat_div = wait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "e2-breadcrumb")))
                cat = wait(cat_div, 5).until(EC.presence_of_all_elements_located((By.TAG_NAME, "span")))[-2].get_attribute('textContent').strip()
                prod['Product Category'] = cat
            except:
                pass

            # scraping product rating
            rating = ''
            try:
                rating = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.score"))).get_attribute('textContent')
                rating = float(rating)
                if rating > 0:
                    prod['Product Rating'] = rating
                else:
                    prod['Product Rating'] = ''
            except:
                pass           
            
            # scraping Image URL link
            url = ''
            try:
                img_div = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.product-gallery")))
                img = wait(img_div, 5).until(EC.presence_of_all_elements_located((By.TAG_NAME, "img")))[0]
                url = img.get_attribute("src")
                if url[:6].lower() == 'https:':
                    prod['Image URL'] = url
                else:
                    prod['Image URL'] = 'https:' + url
            except:
                pass      
            
            # scraping product offers
            offer = ''
            try:
                offer = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.offer"))).get_attribute('textContent').strip()
                prod['Promotion'] = offer
            except:
               pass

            # scraping the product availability
            stock = ''
            try:
                stock = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*=Stock]"))).get_attribute('textContent').strip()
                prod['Availability'] = stock
            except:
                pass

            prod['Input Link'] = input_links[i]
            prod['Extraction Date'] = stamp
               
            # scraping product comments
            if settings['Scrape Comments'] != 0 and prod['Product ID'] != '' and prod['Product Title'] != '' and prod['Price (HKD)'] != '':
                revs_limit = settings['Comment Limit']
                try:
                    rev_div = wait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.reviews-group")))
                    revs = wait(rev_div, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.review")))

                    # applying the comments limit
                    nrevs = len(revs)
                    if nrevs > revs_limit:
                        nrevs = revs_limit

                    for k in range(nrevs):
                        try:
                            comm = {}
                            try:
                                rev = wait(revs[k], 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.review-detail"))).get_attribute('textContent')
                            except:
                                rev = ''
                            try:
                                date_content = wait(revs[k], 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.review-date"))).get_attribute('textContent')
                                elems = re.findall(r'[0-9]+', date_content)
                                elems = elems[::-1]
                                date = '_'.join(elems)
                            except:
                                date = ''

                            try:
                                stars = wait(revs[k], 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "i[class='icon-star active']")))
                                rating = len(stars)
                            except:
                                rating = ''

                            comm['Product ID'] = ID
                            comm['Comment Content'] = rev
                            comm['Comment Rating'] = rating
                            comm['Comment Date'] = date
                            comm['Extraction Date'] = stamp
                            comments = comments.append([comm.copy()]) 
                        except:
                            pass
                except:
                    # No product reviews are available
                    pass

            # checking if the produc data has been scraped successfully
            if prod['Product ID'] != '' and prod['Product Title'] != '' and prod['Price (HKD)'] != '':
                # output scraped data
                prods = prods.append([prod.copy()])
                
        except Exception as err:
            print(f'The error below ocurred during scraping link {i+1}/{nlinks}, skipping ...\n') 
            print(err)
            print('-'*100)
            continue 
        
    # output data
    if prods.shape[0] > 0:
        prods['Extraction Date'] = pd.to_datetime(prods['Extraction Date'], errors='coerce', format="%d_%m_%Y")
        prods['Extraction Date'] = prods['Extraction Date'].dt.date   
        writer = pd.ExcelWriter(output1, date_format='d/m/yyyy')
        prods.to_excel(writer, index=False)
        writer.close()
    if comments.shape[0] > 0:
        comments['Extraction Date'] = pd.to_datetime(comments['Extraction Date'], errors='coerce', format="%d_%m_%Y")
        comments['Extraction Date'] = comments['Extraction Date'].dt.date
        comments['Comment Date'] = pd.to_datetime(comments['Comment Date'], errors='coerce', format="%d_%m_%Y")
        comments['Comment Date'] = comments['Comment Date'].dt.date
        writer = pd.ExcelWriter(output2, date_format='d/m/yyyy')
        comments.to_excel(writer, index=False)
        writer.close()

def initialize_output():

    # removing the previous output file
    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\scraped_data\\' + stamp
    if os.path.exists(path):
        shutil.rmtree(path)

    os.makedirs(path)

    file1 = f'PNS_{stamp}.xlsx'
    file2 = f'PNS_Comments_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
        output2 = path.replace('\\', '/') + "/" + file2

    else:
        output1 = path + "\\" + file1
        output2 = path + "\\" + file2


    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    
    workbook2 = xlsxwriter.Workbook(output2)
    workbook2.add_worksheet()
    workbook2.close()    

    return output1, output2

def get_inputs():
 
    print('Processing The Settings Sheet ...')
    print('-'*100)
    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\PNS_settings.xlsx'
    else:
        path += '/PNS_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "PNS_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        links = []
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Product Link':
                    links.append(row[col])                
                elif col == 'Search Link':
                    links.append(row[col])
                else:
                    settings[col] = row[col]

    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    # checking the settings dictionary
    keys = ["Scrape Comments", "Comment Limit", "Product Limit"]
    for key in keys:
        if key not in settings.keys():
            print(f"Warning: the setting '{key}' is not present in the settings file")
            settings[key] = 0
        try:
            settings[key] = int(float(settings[key]))
        except:
            input(f"Error: Incorrect value for '{key}', values must be numeric only, press an key to exit.")
            sys.exit(1)

    return settings, links

if __name__ == '__main__':

    start = time.time()  
    settings, links = get_inputs()
    output1, output2 = initialize_output()  
    while True:
        try:
            driver = initialize_bot()
            prod_links, input_links = process_links(driver, links, settings)
            scrape_prods(driver, prod_links, input_links, output1, output2, settings)
            driver.quit()
            break
        except Exception as err:
            print('The below error occurred:\n')
            print(err)
            driver.quit()
            time.sleep(5)

    print('-'*100)
    elapsed = round(((time.time() - start)/60), 5)
    hrs = round(elapsed/60, 5)
    input(f'Process is completed successfully in {elapsed} mins ({hrs} hours). Press any key to exit.')
    sys.exit()

