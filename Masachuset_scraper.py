from selenium import webdriver 
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
import os
from selenium.webdriver.chrome.options import Options

excel_file_path = 'Masachusetts_Companies.xlsx'
df = pd.read_excel(excel_file_path)
data_list=df.values.tolist()
count=data_list[0][0] +1

current_dir = os.getcwd()
chrome_driver_path = os.path.join(current_dir, "/chromedriver.exe")
chrome_options = Options()
chrome_options.add_argument(f"webdriver.chrome.driver={chrome_driver_path}")
# Initialize Chrome WebDriver with options
driver = webdriver.Chrome(options=chrome_options)

while True:
    try:
        driver.get(f'https://services.oca.state.ma.us/hic/licdetails.aspx?txtSearchLN={count}')
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/div[1]/div/div/table/tbody/tr/td/table/tbody')))
        tbody=driver.find_element(By.XPATH,'/html/body/form/div[3]/div[1]/div/div/table/tbody/tr/td/table/tbody')
        
        trs=tbody.find_elements(By.TAG_NAME,'tr')
        #------ Registration Number ---------
        td1=trs[0].find_elements(By.TAG_NAME,'td')
        if td1[1].text:
            Registration_number=td1[1].text
        #------ Registrant ---------
        td2=trs[1].find_elements(By.TAG_NAME,'td')
        if td2[1].text:
            Registrant=td2[1].text
        #------ Name ---------
        td3=trs[2].find_elements(By.TAG_NAME,'td')
        if td3[1].text:
            Name=td3[1].text
        #------ Address ---------
        td4=trs[3].find_elements(By.TAG_NAME,'td')
        if td4[1].text:
            Address=td4[1].text
        #------- city -----------
        td5=trs[4].find_elements(By.TAG_NAME,'td')
        if td5[1].text:
            city=td5[1].text
        #-------- Expiration date ----------
        td6=trs[5].find_elements(By.TAG_NAME,'td')
        if td6[1].text:
            Expiration_Date=td6[1].text

        data_list.append([Registration_number, Name, Registrant, Address, city, Expiration_Date])


        count+=1
    except:
        break
        pass

# Convert the first column to integers for sorting
converted_data_list = [[int(x[0])] + x[1:] for x in data_list]

# Sort data_list based on the first column (now all elements are integers)
sorted_data_list = sorted(converted_data_list, key=lambda x: x[0], reverse=True)
# Create a pandas DataFrame from the list of scraped data
df = pd.DataFrame(sorted_data_list, columns=["Registration #", "Name", "Registrant", "Address", "City, State Zip", "Expiration Date"])

# Assuming your data is stored in a DataFrame df
df.replace('<br/>', '', regex=True, inplace=True)

# Create a new Excel workbook
excel_writer = pd.ExcelWriter('Masachusetts_Companies.xlsx', engine='xlsxwriter')

# Convert the DataFrame to an XlsxWriter Excel object
df.to_excel(excel_writer, sheet_name='Sheet1', index=False)

# Get the XlsxWriter workbook and worksheet objects
workbook = excel_writer.book
worksheet = excel_writer.sheets['Sheet1']

# Add some formatting to the Excel file (optional)
header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center', 'fg_color': '#D7E4BC'})
worksheet.set_row(0, None, header_format)

# Save the Excel file
excel_writer._save()