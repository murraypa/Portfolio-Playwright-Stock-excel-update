import selenium
from selenium.webdriver.common.keys import Keys
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

#loop through all rows, skipping title row 1
for symbol in range(2,number_rows+1):
    ticker = sheet['A'+ str(symbol)].value
    print(f"ticker={ticker}")

