import os
import time
import requests
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException

def setup_driver():
    custom_profile_path = os.path.join(os.getcwd(), "chrome_profile")
    if not os.path.exists(custom_profile_path):
        os.makedirs(custom_profile_path)
    chrome_options = Options()
    chrome_options.add_argument(f"--user-data-dir={custom_profile_path}")
    chrome_options.add_experimental_option("prefs", {"profile.default_content_setting_values.cookies": 1})
    chrome_options.add_argument('--ignore-certificate-errors')
    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def get_links(driver):
    driver.get("https://www.mubasher.info/countries/eg/companies")
    driver.implicitly_wait(5)
    tbody = driver.find_element(By.CLASS_NAME, "mi-table")
    names_links_dict = {}
    time.sleep(0.25)
    for table_row in tbody.find_elements(By.XPATH, './/tr'):
        cells = table_row.find_elements(By.XPATH, './/td')
        if len(cells) >= 2:
            driver.implicitly_wait(2.5)
            time.sleep(1)
            name_elmnt = cells[0].text
            link_element = cells[1].find_element(By.XPATH, './/a')
            link = link_element.get_attribute('href')
            modified_link = f"{link}/financial-statements"
            names_links_dict[name_elmnt] = modified_link
    return names_links_dict

def get_financial_data(driver, names_links_dict, selected_companies, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
        csvwriter = csv.writer(csvfile)
        for name, link in names_links_dict.items():
            if name in selected_companies:
                driver.get(link)
                find_selector = driver.find_element(By.TAG_NAME, 'select')
                Select_value = Select(find_selector)
                Select_value.select_by_index(4)
                driver.implicitly_wait(10)
                try:
                    tbody = driver.find_element(By.TAG_NAME, 'table')
                    time.sleep(2)
                    print(f"Table found for {name}: {link}")
                except NoSuchElementException:
                    print(f"Table not found for {name}: {link}")
                    continue
                header_row = tbody.find_element(By.XPATH, ".//tr[1]")
                header_data = [cell.text for cell in header_row.find_elements(By.XPATH, ".//th")]
                csvwriter.writerow([name] + header_data)
                for table_row in tbody.find_elements(By.XPATH, './/tr[position()>1]'):
                    cells = table_row.find_elements(By.XPATH, './/td')
                    if len(cells) >= 5:
                        csvwriter.writerow([name] + [cell.text for cell in cells])

def download_pdfs(driver, names_links_dict, selected_companies):
    if not os.path.exists("pdfs"):
        os.makedirs("pdfs")
    for index, (name, link) in enumerate(names_links_dict.items()):
        if name in selected_companies:
            driver.get(link)
            find_selector = driver.find_element(By.TAG_NAME, 'select')
            Select_value = Select(find_selector)
            Select_value.select_by_index(4)
            driver.implicitly_wait(5)
            time.sleep(0.25)
            pdf_table = driver.find_element(By.CLASS_NAME, "mi-table")
            for table_head in pdf_table.find_elements(By.XPATH, ".//th"):
                try:
                    pdf_link = table_head.find_element(By.XPATH, ".//a").get_attribute('href')
                    response = requests.get(pdf_link)
                    filename = f"pdfs/{name}_pdf_{index + 1}.pdf"
                    with open(filename, 'wb') as f:
                        f.write(response.content)
                        print(f"PDF {index + 1} downloaded successfully for {name}.")
                except NoSuchElementException:
                    continue

def extract_text_from_pdf(pdf_path, language='ara'):
    images = convert_from_path(pdf_path)
    text = ""
    for image in images:
        text += pytesseract.image_to_string(image, lang=language)
    return text

def convert_pdf_to_excel(pdfs, output_dir):
    for pdf_link, company, year in pdfs:
        company_dir = os.path.join(output_dir, company)
        pdf_path = os.path.join(company_dir, f"{year}.pdf")
        text = extract_text_from_pdf(pdf_path)
        data = {'Company': [company], 'Year': [year], 'Text': [text]}
        df = pd.DataFrame(data)
        excel_path = os.path.join(company_dir, f"{year}.xlsx")
        df.to_excel(excel_path, index=False)
        print(f"Converted {pdf_path} to {excel_path}")

