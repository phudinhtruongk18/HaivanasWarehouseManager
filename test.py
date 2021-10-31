import openpyxl as excel
from ExcelManager import read_excel_file
from ObjCanThiet import WareHouse,ProductInDay,ProductOnHand

print("Starting")
raw_data = read_excel_file("./Data/STOCK-ON-HAND.xlsx","SOH GOOD")

print("Got data")

WareHouse(products_on_hand=raw_data)

