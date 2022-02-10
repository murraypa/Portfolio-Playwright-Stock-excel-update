from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import openpyxl
from openpyxl import Workbook

print('open xl file')
#open excel workbook
wb = openpyxl.load_workbook('./Book.xlsx')
#get worksheet names
print(f"worksheet names: {wb.sheetnames}")
#get single worksheet
sheet = wb["ticker symbols"]
print(f"sheet={sheet}")
#set active worksheet
ws = wb.active
ticker = sheet['A2'].value
print(f"row1={ticker}\n")
number_rows = ws.max_row
print(f"max rows in sheet:{number_rows}")

#control web browser with web driver
#driver = webdriver.Chrome(".\chromedriver.exe")
#driver = webdriver.chrome()

driver = Service(r'C:\\Users\\user\\Downloads\\python\\Selenium\\selenium-tests\\geckodriver.exe')
#driver = webdriver.Firefox(executable_path=r'C:\Users\user\Downloads\python\Selenium\selenium-tests\geckodriver.exe')
driver.get("https://finance.yahoo.com/")
quotelookup = driver.find_element(By.CSS_SELECTOR,'#darla_csc_holder').send_keys(ticker)
quotelookup = driver.find_element(By.CSS_SELECTOR,'#darla_csc_holder').send_keys(Keys.ENTER)
#quotelookup = driver.find_element_by_css_selector('#darla_csc_holder').send_keys(ticker) #  #input.D\(ib\)
#quotelookup = driver.find_element_by_css_selector('#darla_csc_holder').send_keys(Keys.ENTER)  ##darla_csc_holder #input.D\(ib\)


#loop through all rows, skipping title row 1
for symbol in range(2,number_rows+1):
    ticker = sheet['A'+ str(symbol)].value
    print(f"ticker={ticker}")
    #put current ticker into "quote lookup" box
    quotelookup = driver.find_element_by_css_selector('input.D\(ib\)').send_keys(ticker)
    quotelookup = driver.find_element_by_css_selector('input.D\(ib\)').send_keys(Keys.ENTER)

