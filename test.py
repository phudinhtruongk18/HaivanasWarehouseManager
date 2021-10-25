import openpyxl as excel
from ExcelManager import read_excel_file

data = read_excel_file("./Data/STOCK-ON-HAND.xlsx","SOH GOOD")

print(data)

