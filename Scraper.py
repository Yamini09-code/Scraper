import os
import time
import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager as CM


def save_to_excel(file_name, prompt, answer, links):

    if os.path.exists(file_name):

        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active

        sheet.append(["Prompt", "Answer", "Link"])

    
    first_row = True
    for link in links:
        
        if first_row:
            sheet.append([prompt, answer, link])
            first_row = False
        else:
            
            sheet.append(["", "", link])

 
    workbook.save(file_name)
    print(f"Data saved to {file_name}")


chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
service = Service(executable_path=CM().install())

driver = webdriver.Chrome(service=service, options=chrome_options)

url_link = "https://www.perplexity.ai/"
driver.get(url_link)

time.sleep(8)
driver.implicitly_wait(5)


search_box = driver.find_element(By.TAG_NAME, "textarea")
query = "machine learning"
search_box.send_keys(query)
search_box.send_keys(Keys.RETURN)


time.sleep(5)


data = []
try:

    text_results = driver.find_elements(
        By.XPATH,
        "/html/body/main/div/div/div[2]/div/div/div[3]/div/div[1]/div[1]/div/div/div[3]/div/div[1]/div[2]/div[2]",
    )
    result_text = text_results[0].text if text_results else "No result found"
    

    list = driver.find_elements(
        By.XPATH,
        "//main/div/div/div[2]/div/div/div[3]/div/div[1]/div[1]/div/div/div[3]/div/div[1]/div[2]/div[1]/div/div/div[2]/div/div",
    )

    list[-1].click()

    links = driver.find_elements(
        By.XPATH, "/html/body/div/div[2]/div/div/div/div[2]/div[2]/div/div/div[1]/a"
    )

    link_urls = [i.get_attribute("href") for i in links]

    driver.implicitly_wait(5)


finally:

    driver.quit()


file_name = "perplexity_results.xlsx"
save_to_excel(file_name, query, result_text, link_urls)
