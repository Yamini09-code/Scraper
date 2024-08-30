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

from Commons import save_to_excel, read_csv_to_dataframe, get_main_content

user_input = read_csv_to_dataframe("Input.csv")  

prompt = user_input.loc[user_input['Input_Variables'] == "Prompt", 'Value'].values[0]
new_run = user_input.loc[user_input['Input_Variables'] == "New_Run", 'Value'].values[0]

new_run = new_run == "TRUE"


chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
service = Service(executable_path=CM().install())

driver = webdriver.Chrome(service=service, options=chrome_options)

url_link = "https://www.perplexity.ai/"
driver.get(url_link)

time.sleep(8)
driver.implicitly_wait(5)


search_box = driver.find_element(By.TAG_NAME, "textarea")
search_box.send_keys(prompt)
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
answer = result_text

text_list = []
valid_links = []

for link in link_urls:
    text = get_main_content(link)
    if text:  # Only add if the text is successfully retrieved
        text_list.append(text)
        valid_links.append(link)

data = []

# Only append if there are valid links and corresponding text
if valid_links and text_list:
    data.append([prompt, answer, valid_links[0], text_list[0]])

for i in range(1, len(valid_links)):
    data.append(["", "", valid_links[i], text_list[i]])

df = pd.DataFrame(data, columns=["Prompt", "Answer", "Link", "Text"])
save_to_excel(file_name, df, new_run)

