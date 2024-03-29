from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
from openpyxl import Workbook
import requests
import time
from selenium.webdriver.chrome.options import Options
import os
import re
from urllib.parse import urlparse, parse_qs

def extract_ctr(link):
    parsed_url = urlparse(link)
    query_params = parse_qs(parsed_url.query)
    ctr_number = query_params.get('ctr', [''])[0]
    return int(ctr_number)

excel_file_path = "Kentucky_excel_file.xlsx"

df2 = pd.read_excel(excel_file_path)
data = df2.values.tolist()
match = re.search(r'\d+$', data[0][0])
count = int(match.group())+1
current_dir = os.getcwd()

chrome_driver_path = os.path.join(current_dir, "/chromedriver.exe")
chrome_options = Options()
chrome_options.add_argument(f"webdriver.chrome.driver={chrome_driver_path}")
# Initialize Chrome WebDriver with options
driver = webdriver.Chrome(options=chrome_options)
#action = ActionChains(driver)
# Excel filename
excel_filename = "Kentucky_excel_file.xlsx"
columns = ["URL", "Name", "Status", "File Date", "Principal Office", "Registered Agent"] 

wb = Workbook()
ws = wb.active
header = columns
ws.append(header)


error=0
data_count=0
while True:
    try:
        url="https://web.sos.ky.gov/bussearchnprofile/Profile.aspx/?ctr="+str(count)
        driver.get(url)
        response=requests.head(url)
        if error==3:
            break
        if response.status_code==500 or response.status_code==404:
            count+=1
            error+=1
            continue
        error=0
        try: 
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'Activebg')))
        except:
            # Open a new tab using JavaScript
            driver.execute_script("window.open('', '_blank');")
            # Switch to the new tab
            driver.switch_to.window(driver.window_handles[1])
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(.5)
        content = driver.page_source
        soup=BeautifulSoup(content,'html.parser')
        # Find the table using the specific style attribute

        root_tr_elements = soup.find_all("tr",class_='Activebg')

        if len(root_tr_elements)==14:
            status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
            if status=='A - Active':    
                Name=root_tr_elements[1].find('td',style='font-family:Book Antiqua;width:67%;').text
                status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
                File_Date=root_tr_elements[8].find('td',style='font-family:Book Antiqua;width:67%;').text
                if root_tr_elements[8].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='File Date':
                    File_Date=root_tr_elements[9].find('td',style='font-family:Book Antiqua;width:67%;').text 
                principal_office=root_tr_elements[11].find('td',style='font-family:Book Antiqua;width:67%;').text
                if root_tr_elements[11].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='Principal Office':
                    principal_office=root_tr_elements[12].find('td',style='font-family:Book Antiqua;width:67%;').text 
                try:
                    if root_tr_elements[13].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text=='Registered Agent':
                        registered_Agent=root_tr_elements[13].find('td',style='font-family:Book Antiqua;width:67%;')
                        registered_Agent=registered_Agent.next
                    elif root_tr_elements[13].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='Registered Agent':
                        registered_Agent=root_tr_elements[12].find('td',style='font-family:Book Antiqua;width:67%;')
                        registered_Agent=registered_Agent.next
                except:
                    registered_Agent=root_tr_elements[13].find('td',style='font-family:Book Antiqua;width:67%;').text   
                    if root_tr_elements[13].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='Registered Agent':
                        registered_Agent=root_tr_elements[12].find('td',style='font-family:Book Antiqua;width:67%;').text
                data.append([url, Name, status, File_Date, principal_office, registered_Agent])  
            else:
                count+=1
                continue
        elif len(root_tr_elements)==11:
            status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
            if status=='A - Active':
                Name=root_tr_elements[1].find('td',style='font-family:Book Antiqua;width:67%;').text
                status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
                File_Date=root_tr_elements[7].find('td',style='font-family:Book Antiqua;width:67%;').text
                principal_office=root_tr_elements[10].find('td',style='font-family:Book Antiqua;width:67%;').text
                registered_Agent=root_tr_elements[11].find('td',style='font-family:Book Antiqua;width:67%;')
                try:
                    registered_Agent=registered_Agent.next
                except:
                    registered_Agent=registered_Agent.text   
                data.append([url, Name, status, File_Date, principal_office,registered_Agent])
            else:
                count+=1
                continue
        elif len(root_tr_elements)==12:
            status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
            if status=='A - Active':
                Name=root_tr_elements[1].find('td',style='font-family:Book Antiqua;width:67%;').text
                status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
                File_Date=root_tr_elements[7].find('td',style='font-family:Book Antiqua;width:67%;').text
                principal_office=root_tr_elements[10].find('td',style='font-family:Book Antiqua;width:67%;').text
                registered_Agent=root_tr_elements[11].find('td',style='font-family:Book Antiqua;width:67%;')  
                try:
                    registered_Agent=registered_Agent.next
                except:
                    registered_Agent=registered_Agent.text   
                data.append([url, Name, status, File_Date, principal_office,registered_Agent])
            else:
                count+=1
                continue
        elif len(root_tr_elements)==13:
            status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
            if status=='A - Active':
                Name=root_tr_elements[1].find('td',style='font-family:Book Antiqua;width:67%;').text
                status=root_tr_elements[4].find('td',style='font-family:Book Antiqua;width:67%;').text
                File_Date=root_tr_elements[7].find('td',style='font-family:Book Antiqua;width:67%;').text
                if root_tr_elements[7].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='File Date':
                    File_Date=root_tr_elements[8].find('td',style='font-family:Book Antiqua;width:67%;').text 
                principal_office=root_tr_elements[10].find('td',style='font-family:Book Antiqua;width:67%;').text
                if root_tr_elements[10].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='Principal Office':
                    principal_office=root_tr_elements[11].find('td',style='font-family:Book Antiqua;width:67%;').text 
                try:
                    if root_tr_elements[12].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text=='Registered Agent':
                        registered_Agent=root_tr_elements[12].find('td',style='font-family:Book Antiqua;width:67%;')
                        registered_Agent=registered_Agent.next
                    elif root_tr_elements[12].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='Registered Agent':
                        registered_Agent=root_tr_elements[11].find('td',style='font-family:Book Antiqua;width:67%;')
                        registered_Agent=registered_Agent.next
                except:
                    registered_Agent=root_tr_elements[12].find('td',style='font-family:Book Antiqua;width:67%;').text   
                    if root_tr_elements[12].find('td',style='font-family:Arial;font-weight:bold;width:25%;').text!='Registered Agent':
                        registered_Agent=root_tr_elements[11].find('td',style='font-family:Book Antiqua;width:67%;').text 
                data.append([url, Name, status, File_Date, principal_office,registered_Agent])
            else:
                count+=1
                continue
        elif len(root_tr_elements)==15:
            status=root_tr_elements[5].find('td',style='font-family:Book Antiqua;width:67%;').text
            if status=='A - Active':
                Name=root_tr_elements[1].find('td',style='font-family:Book Antiqua;width:67%;').text
                status=root_tr_elements[5].find('td',style='font-family:Book Antiqua;width:67%;').text
                File_Date=root_tr_elements[9].find('td',style='font-family:Book Antiqua;width:67%;').text
                principal_office=root_tr_elements[13].find('td',style='font-family:Book Antiqua;width:67%;').text
                registered_Agent=root_tr_elements[14].find('td',style='font-family:Book Antiqua;width:67%;')
                try:
                    registered_Agent=registered_Agent.next
                except:
                    registered_Agent=registered_Agent.text     
                data.append([url, Name, status, File_Date, principal_office,registered_Agent])
            else:
                count+=1
                continue
        count+=1
        data_count+=1
        if data_count==50:
            try:
                df = pd.DataFrame(data,columns=columns)
                # Preprocess each element in the data list to remove <br/> tags
                cleaned_data = []
                for row in data:
                    cleaned_row = []
                    for cell in row:
                        # Remove or replace <br/> tags as needed
                        cleaned_cell = str(cell).replace("<br/>", "\n")
                        cleaned_row.append(cleaned_cell)
                    cleaned_data.append(cleaned_row)

                # Append data to the Excel file
                for data_row in cleaned_data:
                    ws.append(data_row)

                # Save the Excel file
                wb.save(excel_filename)
                data.clear()
                data_count=0
            except Exception as e:
                print("error occured",e)
                data_count=0
    except:
        count+=1
        

driver.quit()
if data:
    # Create a pandas DataFrame from the collected data
    df = pd.DataFrame(data, columns=columns)
    cleaned_data = []
    for row in data:
        cleaned_row = []
        for cell in row:
            # Remove or replace <br/> tags as needed
            cleaned_cell = str(cell).replace("<br/>", "\n")
            cleaned_row.append(cleaned_cell)
        cleaned_data.append(cleaned_row)
    for data_row in cleaned_data:
        ws.append(data_row)
    # Save the Excel file
    wb.save(excel_filename)

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(excel_file_path)

df['ctr_number'] = df.iloc[:, 0].apply(extract_ctr)

# Sort the DataFrame based on the first column
#df_sorted = df.sort_values(by=df.columns[0], ascending=False)

df_sorted = df.sort_values(by='ctr_number', ascending=False)
# Drop the temporary column
df_sorted.drop(columns=['ctr_number'], inplace=True)
df_sorted.to_excel(excel_file_path, index=False)


# Read the Excel file into a DataFrame
df = pd.read_excel("Kentucky_excel_file.xlsx")

# Remove duplicate rows
df.drop_duplicates(inplace=True)

# Write the DataFrame back to Excel
df.to_excel("Kentucky_excel_file.xlsx", index=False)