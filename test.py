from ExcelManager import read_warehouse
from ObjCanThiet import WareHouse,ProductInDay,ProductOnHand,SaleData

#
# print("Starting")
# raw_data = read_warehouse("./Data/STOCK-ON-HAND.xlsx","SOH GOOD")
# print("Got data")
# WareHouse(products_on_hand=raw_data)
#

stock_in_day = SaleData("./Data/day_sale.xlsx")
print(stock_in_day.sheet_names)

get_sheet = "31.10"

in_stocks = stock_in_day.get_list_product(get_sheet)

