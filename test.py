from ExcelManager import read_warehouse
from ObjCanThiet import WareHouse,ProductInDay,ProductOnHand,SaleData

stock_in_day = SaleData("./Data/day_sale.xlsx")
print(stock_in_day.sheet_names)

get_sheet = "31.10"
stock_in_day.get_list_product(get_sheet)
print(stock_in_day.products_in_day.__len__())

raw_data = read_warehouse("./Data/STOCK-ON-HAND.xlsx","SOH GOOD")
warehouse = WareHouse(products_on_hand=raw_data)

# they are romeo and juliet in 2021
warehouse.set_sale_in_day(stock_in_day.products_in_day)
# check stock not in warehouse
stock_non_def = warehouse.get_stocks_not_in_warehouse()
print(stock_non_def)
if stock_non_def.__len__() > 0:
    print("Hien thong bao va xuat file text")
    warehouse.export_text_file_non_define_stock_in_warehouse(stock_non_def)
    # return ngay day

warehouse.cuculate_stock()
warehouse.export()
