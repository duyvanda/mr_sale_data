import os
from google.oauth2 import service_account
from googleapiclient import discovery
import pandas as pd
import numpy as np
import xlwings as xw
import unidecode

# DMS download
from selenium import webdriver
import time
PATH = "C:/Users/DELL/selenium/chromedriver_win32/chromedriver.exe"
driver = webdriver.Chrome(PATH)
driver.get('http://dms.phanam.com.vn/')
driver.maximize_window()
user_id = "txtUserName-inputEl"
password_id = "txtPassword-inputEl"
dangnhap_id = "btnLogin-btnIconEl"
driver.find_element_by_id(user_id).send_keys("MR2523")
driver.find_element_by_id(password_id).send_keys("duyvq@123")
driver.find_element_by_id(dangnhap_id).click()
time.sleep(2)
classname = 'x-tree-node-text '
driver.find_elements_by_class_name(classname)[0].click()
time.sleep(2)
driver.find_elements_by_class_name(classname)[1].click()
time.sleep(2)
driver.find_elements_by_class_name(classname)[4].click()
time.sleep(2)
seq = driver.find_elements_by_tag_name('iframe')
print("No of frames present in the web page are: ", len(seq))
iframe = driver.find_elements_by_tag_name('iframe')[0]
driver.switch_to.frame(iframe)
tungayid = 'cboDate00-inputEl'
driver.find_element_by_id(tungayid).clear()
driver.find_element_by_id(tungayid).send_keys("01-06-2021")
time.sleep(1)
macongtyid = 'chkList0-inputEl'
driver.find_element_by_id(macongtyid).click()
time.sleep(1)
tabkieudonhangid = 'tab-1131-btnEl'
driver.find_element_by_id(tabkieudonhangid).click()
time.sleep(1)
checkkieudonhangid = 'chkList1-inputEl'
driver.find_element_by_id(checkkieudonhangid).click()
time.sleep(1)
exportid = 'btnExport-btnInnerEl'
driver.find_element_by_id(exportid).click()
time.sleep(60)
driver.switch_to.default_content()
driver.close()

# Excel Handling
app = xw.App(visible=False)
wb = app.books.open('C:/Users/DELL/Downloads/Rawdata Doanh Số Chi Tiết (Tính Lương).Xlsb')
wb.sheets['Data'].range('1:4').delete()
wb.save()
wb.close()
app.quit()

# Google Sheet connection

scopes = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/drive.file']
jsonfile = os.path.join(os.getcwd(),'datateam1599968716114-6f9f144b4262.json')
credentials = service_account.Credentials.from_service_account_file(jsonfile, scopes = scopes)
service = discovery.build('sheets','v4',credentials = credentials)
def datetime_to_int(dt):
    return int(dt.strftime("%Y%m%d%H%M%S"))
df = pd.read_excel("C:/Users/DELL/Downloads/Rawdata Doanh Số Chi Tiết (Tính Lương).Xlsb", engine='pyxlsb')# , engine='pyxlsb'
df.replace(np.nan, '', inplace=True)
columns = df.columns
columns_list = []
for c in columns:
    c = unidecode.unidecode(c)
    columns_list.append(c)

df.columns = columns_list
df.columns = df.columns.str.replace(' ','_')
spreadsheets_id ='1qEwviiJcAtvWCLvc-5AXOgidaeHMsi2RWWIKiUehabI'
rangeAll = '{0}!A1:BI'.format('DF1')
body = {}
resultClear = service.spreadsheets().values().clear( spreadsheetId=spreadsheets_id, range=rangeAll,body=body ).execute()
response_date = service.spreadsheets().values().append(
    spreadsheetId=spreadsheets_id,
    valueInputOption='RAW',
    range='DF1!A1',
    body=dict(
        majorDimension='ROWS',
        values=df.T.reset_index().T.values.tolist())
).execute()
os.remove("C:/Users/DELL/Downloads/Rawdata Doanh Số Chi Tiết (Tính Lương).Xlsb")