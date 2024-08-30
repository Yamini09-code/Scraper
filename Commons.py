
import openpyxl
import os
import pandas as pd
import requests
from bs4 import BeautifulSoup

def save_to_excel(file_name, df,new_run = False):   

    if new_run and os.path.exists(file_name):
        os.remove(file_name)   


    if os.path.exists(file_name):  

        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
 
          
        sheet.append(df.columns.tolist())

        
    for _, row in df.iterrows():
        sheet.append(row.tolist())

    workbook.save(file_name)
    print(f"Data saved to {file_name}")   




def read_csv_to_dataframe(file_path):
    try:
        df = pd.read_csv(file_path)
        return df
    except FileNotFoundError:
        print(f"The file {file_path} was not found.")
    except pd.errors.EmptyDataError:
        print("No data found in the file.")
    except pd.errors.ParserError:
        print("Error parsing the file.")
    except Exception as e:
        print(f"An error occurred: {e}")



def get_main_content(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            main_content = soup.find_all(['article', 'p'])
            text = ' '.join([content.get_text(separator=' ') for content in main_content])
            return text.strip()
        else:
            return None  # Indicate failure by returning None
    except requests.exceptions.RequestException:
        return None  # Indicate failure by returning None












