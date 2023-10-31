from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import itertools
import pandas as pd
import requests
from datetime import datetime
from datetime import date
from tkinter import filedialog

# Global Variables
waybill = []
result = [] # For normal waybills
result1 = [] # For abnormal waybills
columns = ['Forwarder', 'Suppliers', 'PO', 'Master Air Waybill No.', 'House Air Waybill No.', 'ETA', 'Freight No.', 'Port', 'Country', 'No. of Case']

def setup(waybill_list):
    waybill.extend(waybill_list)

    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    for line in waybill:
        if line[0] in ["UPS", "DHL", "SF"]:
            driver = webdriver.Chrome("chromedriver.exe",options=options)
            driver.minimize_window()
            return driver
    return None
     

def case_ups(driver,waybillno):
    temp = ['UPS']
    wait = WebDriverWait(driver,10)
    driver.get("http://upsfd.ecfreight.net/")

    driver.find_element(By.TAG_NAME,'textarea').send_keys(waybillno)
    driver.find_element(By.CLASS_NAME,"MyButton").click()

    try:
        wait.until(EC.presence_of_element_located((By.XPATH,"//*[contains(text(), '{}')]".format(waybillno))))
    except TimeoutException:
        return True
        
    soup = BeautifulSoup(driver.page_source,'html.parser')
    soup = soup.find(class_="gv_RowStyle")
    
    for item in soup.find_all('td'):
        temp.append(item.text.strip())

    mapping = {
    'Forwarder': 0,
    'Suppliers': None,
    'PO': None,
    'Master Air Waybill No.': 7,
    'House Air Waybill No.': 2,
    'ETA': 1,
    'Freight No.': 4,
    'Port': 6,
    'Country': None,
    'No. of Case': None
    }
    new_row = []
    if temp[1]!= "Shipment Date":
        
        for column in columns:
            index = mapping[column]
            if column == 'ETA':
                new_row.append(datetime.strptime(temp[index], '%Y-%m-%d').strftime('%d/%m/%Y'))
            elif column == 'Master Air Waybill No.':
                if temp[5] == "Air": 
                    new_row.append(temp[index])
                elif temp[5] == "Road":
                    new_row.append("CCRN:{}".format(temp[index]))
            elif column == "Country":
                new_row.append("CN") if temp[5] == "Road" else new_row.append('')
            elif index is not None:
                new_row.append(temp[index])
            else:
                new_row.append('')
        result.append(new_row)

    else:
        new_row = [temp[0],'','','No information available',waybillno,'','','','','']
        result1.append(new_row)


def case_dhl(driver,waybillno):
    temp = [['DHL']]

    driver.get('https://apps.dhl.com.hk/eng_fi/')
    frame_ele = driver.find_elements(By.TAG_NAME,'frame')[1]
    driver.switch_to.frame(frame_ele)

    driver.find_element(By.NAME,"airbill").send_keys(waybillno)
    driver.find_element(By.CSS_SELECTOR,"input[type=\"submit\" i]").click()

    soup = BeautifulSoup(driver.page_source,'html.parser')
    soup.find('blockquote')
    try:
        soup = soup.find_all('tr')[1]
    except IndexError:
        
        return True
    
    for item in soup.find_all('b'):
        a = item.text.strip()
        temp.append(a.split(', '))
    
    temp = list(itertools.chain(*temp))

    # Define a dictionary that maps the column names to the corresponding indices in the result list

    mapping = {
        'Forwarder': 0,
        'Suppliers': None,
        'PO': None,
        'Master Air Waybill No.': 4,
        'House Air Waybill No.': 1,
        'ETA': 3,
        'Freight No.': 2,
        'Port': 5,
        'Country': len(temp)-2,
        'No. of Case': len(temp)-1
    }

    # Create a new list of lists that contains the data in the correct order and with the correct column headers

    new_row = []
    if temp[2]!= "No information available":
        for column in columns:
            index = mapping[column]
            if column == 'ETA':
                new_row.append(datetime.strptime(temp[index], '%d-%m-%Y').strftime('%d/%m/%Y'))
            elif column == 'Master Air Waybill No.':
                new_row.append(temp[index][:3]+'-'+temp[index][3:])
            elif index is not None:
                new_row.append(temp[index])
            else:
                new_row.append('')
        result.append(new_row)
 
    else:
        new_row = [temp[0],'','','No information available',temp[1],'','','','','']
        result1.append(new_row)


def case_fedex(forwarder,waybillno):
    temp = [[forwarder.upper()]]

    response = requests.get('https://wap-asia.fedex.com/hkodds/en-hk/DeclarationDetails.htm?tracknumbers={}%0D%0A&action=show'.format(waybillno))
    soup = BeautifulSoup(response.content,"html.parser")

    try:
        item = soup.find_all('table')[3]

    except IndexError:
        
        return True
    
    info = item.find_all('tr')[2]
    temp.append(info.text.strip().split('\n'))
    temp = list(itertools.chain(*temp))
    mapping = {
    'Forwarder': 0,
    'Suppliers': None,
    'PO': None,
    'Master Air Waybill No.': None,
    'House Air Waybill No.': 1,
    'ETA': 3,
    'Freight No.': 8,
    'Port': 6,
    'Country': None,
    'No. of Case': None
    }
    new_row = []
    if temp[3] != '':
        for column in columns:
            index = mapping[column]
            if column == 'ETA':
                new_row.append(datetime.strptime(temp[index], '%b %d, %Y').strftime('%d/%m/%Y'))
            elif column == 'Master Air Waybill No.':
                if temp[7] == "Air": 
                    new_row.append(temp[4][:3]+'-'+temp[4][3:])
                elif temp[7] == "Road":
                    new_row.append("CCRN:{}".format(temp[5]))
            elif column == "Country":
                new_row.append("CN") if temp[7] == "Road" else new_row.append('')
            elif index is not None:
                new_row.append(temp[index])
            else:
                new_row.append('')
        result.append(new_row)
    else:
        new_row = [temp[0],'','','No information available',temp[1],'','','','','']
        result1.append(new_row)


def case_sf(driver,waybillno):
    temp = ["SF"] 
    wait = WebDriverWait(driver,30)
    driver.get('https://htm.sf-express.com/hk/en/dynamic_function/more/CcrnSearch/')
    
    driver.find_element(By.CLASS_NAME,"token-input").send_keys(waybillno)
    driver.find_element(By.TAG_NAME,"h2").click()

    try:
        wait.until(EC.presence_of_element_located((By.CLASS_NAME,"primary-button")))
    
    except TimeoutException:
        return True

    driver.find_element(By.CLASS_NAME,"primary-button").click()

    try:
        wait.until(EC.presence_of_element_located((By.CLASS_NAME,"table-bordered")))
    
    except TimeoutException:
        return True

    soup = BeautifulSoup(driver.page_source,"html.parser")
    soup = soup.find(class_="table table-bordered")
    
    try:
        soup = soup.find("tbody")
    except AttributeError:
        return True
    
    for item in soup.find_all("td"):
        temp.append(item.text)

    mapping = {
            'Forwarder': 0,
            'Suppliers': None,
            'PO': None,
            'Master Air Waybill No.': 3 if temp[3]!="-" else 2,
            'House Air Waybill No.': 1,
            'ETA': 6,
            'Freight No.': 4 if temp[4]!="-" else 5,
            'Port': 7,
            'Country': None,
            'No. of Case': None
        }

    new_row = []
    for column in columns:
        index = mapping[column]
        if column == 'Master Air Waybill No.':
            if temp[3]=="-":
                new_row.append("CCRN:{}".format(temp[index]))
            else:
                new_row.append(temp[index][:3]+'-'+temp[index][3:])
        elif column == 'Country':
            if temp[3]=="-":
                new_row.append("CN")
            else:
                new_row.append('')

        elif index is not None:
            new_row.append(temp[index])
        else:
            new_row.append('')
     
    result.append(new_row)


def main_func(driver): 
    for iteration in range(len(waybill)):
        forwarder = waybill[iteration][0]
        if forwarder =='' or waybill[iteration][1]=='':
            continue
        elif forwarder.lower() == 'dhl':
            case_dhl(driver,iteration)
        elif forwarder.lower() == 'ups':
            case_ups(driver,iteration)
        else:
            case_fedex(iteration) 

    for line in waybill:
        if line[0] in ["UPS", "DHL"]:
            driver.quit()
            break


def output_func():
    sorted_result = sorted(result, key=lambda x: datetime.strptime(x[5],'%d/%m/%Y'),reverse=True)

    output = sorted_result + result1
    if len(output) != 0:
        df = pd.DataFrame(output, columns=columns)

        folder_selected = filedialog.askdirectory(title="Please select a location for the excel output")

        # Write the DataFrame to a designated file with headers
        today = date.today()
        formatted_date = today.strftime("%Y%m%d")
        filename = '{}/RPA Custom Declaration on {}.xlsx'.format(folder_selected,formatted_date)

        df.to_excel(filename, index=False)
    
    result.clear()
    result1.clear()
    

 


