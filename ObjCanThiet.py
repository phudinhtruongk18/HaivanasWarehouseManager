from tkinter import messagebox
import openpyxl as excel
from ExcelManager import export_warehouse
import datetime
# import pandas
# import matplotlib.pyplot as plt


class ProductInDay:
    def __init__(self, bar_code):
        self.bar_code = str(bar_code).strip()
        self.sale_quantity = 0
        self.in_quantity = 0

    def set_sale_quantity(self, quantity):
        self.sale_quantity = quantity

    def set_in_quantity(self, quantity):
        self.in_quantity = quantity

    def add_sale_quantity(self, quantity):
        self.sale_quantity += quantity

    def add_in_quantity(self, quantity):
        self.in_quantity += quantity

    def show_info(self):
        print(self.bar_code, str(self.sale_quantity), str(self.in_quantity))

    def is_in_warehouse(self,bar_code_in_warehouse):
        if self.bar_code in bar_code_in_warehouse:
            return True
        else:
            print("not in ware house",self.bar_code)
            return False


class SaleData:
    def __init__(self, fileName):
        self.workbook = excel.load_workbook(filename=fileName)
        self.sheet_names = self.workbook.get_sheet_names()
        self.products_in_day = []
        self.checked_sale = []

    def get_sheets(self):
        return self.sheet_names

    def get_sheet_data(self, sheet_name):
        data = self.workbook[sheet_name]
        return data

    def get_sale_stocks(self, sheet_name):
        data = self.get_sheet_data(sheet_name=sheet_name)
        sale_stocks = []

        for index, data_temp in enumerate(data.values):
            # bo hai hang đau tien
            if index < 2:
                continue
            # ket thuc neu xuat hien chu nay
            if data_temp[1] is None or data_temp[1] == "":
                break
            sale_stocks.append(data_temp[1:3])

        return sale_stocks

    def get_shop_sale_stocks(self, sheetname):
        data = self.get_sheet_data(sheet_name=sheetname)
        sale_stocks = []

        for index, data_temp in enumerate(data.values):
            # bo hai hang đau tien
            if index < 2:
                continue
            # ket thuc neu xuat hien chu nay
            if data_temp[6] is None or data_temp[6] == "":
                break
            sale_stocks.append(data_temp[6:8])
        return sale_stocks

    def get_in_stocks(self, sheetname):
        data = self.get_sheet_data(sheet_name=sheetname)
        in_stocks = []

        for index, data_temp in enumerate(data.values):
            # bo hai hang đau tien
            if index < 2:
                continue
            # ket thuc neu xuat hien chu nay
            if data_temp[11] is None or data_temp[11] == "":
                break
            in_stocks.append(data_temp[11:13])
        return in_stocks

    def find_index_sale_product(self, bar_code):
        for index_product, temp_pro in enumerate(self.products_in_day):
            if temp_pro.bar_code == bar_code:
                return index_product

    def get_list_product(self, sheet_name):
        sale_stock = self.get_sale_stocks(sheet_name)
        shop_sale_stock = self.get_shop_sale_stocks(sheet_name)
        sum_sale_list = sale_stock + shop_sale_stock
        in_stocks = self.get_in_stocks(sheet_name)

        for temp_sale in sum_sale_list:
            bae_code = temp_sale[0]
            quantity = temp_sale[1]
            if bae_code in self.checked_sale:
                # neu checked list thi tim index va plus in quantity
                index_product = self.find_index_sale_product(bae_code)
                self.products_in_day[index_product].add_sale_quantity(quantity)

            else:
                # neu khong trong checked list them vao sale_product
                self.checked_sale.append(bae_code)
                product = ProductInDay(bar_code=bae_code)
                product.set_sale_quantity(quantity=quantity)
                self.products_in_day.append(product)

        # tuong tu voi hang nhap vao
        for temp_sale in in_stocks:
            bae_code = temp_sale[0]
            quantity = temp_sale[1]
            if bae_code in self.checked_sale:
                # neu checked list thi tim index va plus in quantity
                index_product = self.find_index_sale_product(bae_code)
                self.products_in_day[index_product].add_in_quantity(quantity)
            else:
                # neu khong trong checked list them vao sale_product
                self.checked_sale.append(bae_code)
                product = ProductInDay(bar_code=bae_code)
                product.set_in_quantity(quantity=quantity)
                self.products_in_day.append(product)


class ProductOnHand:
    def __init__(self,index, *args):
        self.is_updated = False
        self.bar_code = str(args[0]).strip()
        self.en_name = args[1]
        self.vn_name = args[2]
        # self.in_stock = args[3]
        # self.sale_stock = args[4]
        self.in_stock = 0
        self.sale_stock = 0

        self.ava_stock = args[5]
        if not isinstance(self.ava_stock, int):
            print('not int', self.ava_stock)
            raise TypeError("Warehouse file is wrong at ",self.bar_code + " - row " + str(index+4) +" in warehouse")

        self.new_ava_stock = 0

    def excel_format(self):
        if self.is_updated:
        # not get self.ava_stock just need new_ava_stock after caculate
            return [self.bar_code, self.en_name, self.vn_name, self.in_stock, self.sale_stock, self.new_ava_stock]
        else:
            return [self.bar_code, self.en_name, self.vn_name, self.in_stock, self.sale_stock, self.ava_stock]


    def __str__(self):
        return f"\n {self.bar_code} en_name: {self.en_name} + vn_name {self.vn_name} " \
               f"in_stock {self.in_stock}sale_stock {self.sale_stock} ava_stock {self.ava_stock}" \
               f"  new_ava_stock {self.new_ava_stock}"

    def update_stock(self,stock_in_day):
        self.sale_stock = stock_in_day.sale_quantity
        self.in_stock = stock_in_day.in_quantity
        self.new_ava_stock = self.ava_stock - self.sale_stock + self.in_stock
        self.is_updated = True


    def to_dict(self):
        return {
            'bar_code': self.bar_code,
            'en_name': self.en_name,
            'vn_name': self.vn_name,
            'in_stock': self.in_stock,
            'sale_stock': self.sale_stock,
            'new_ava_stock': self.new_ava_stock,
        }


class WareHouse(list):
    def __init__(self, products_on_hand):
        super().__init__()
        self.bar_code_list = []
        for index,product in enumerate(products_on_hand):
            product = ProductOnHand(index,*product)
            self.append(product)
            self.bar_code_list.append(product.bar_code)
        self.sale_in_day = None

    def set_sale_in_day(self, sale_in_day):
        self.sale_in_day = sale_in_day

    def show_stock(self):
        for product_on_hand in self:
            print(product_on_hand)

    def find_index_stock(self,stock):
        for index,temp_pro in enumerate(self):
            if temp_pro.bar_code == stock.bar_code:
                return index
        # khong co truong hop else duoc boi vi da kiem tra 1 lan

    def cuculate_stock(self):
        # da~ pass bai kiem tra ton tai thu nhat
        # nen gio chi can return index cua no va thay doi
        # in sale new ava de export
        for stock in self.sale_in_day:
            # print("=============================")
            # print("Before")
            index_stock = self.find_index_stock(stock)
            # print(self[index_stock])

            # print("After")
            self[index_stock].update_stock(stock)
            # print(self[index_stock])
            # print("=============================")

    def export(self,sum_in, sum_sale, sum_ava):
        export_warehouse(self,sum_in, sum_sale, sum_ava)

    def get_stocks_not_in_warehouse(self):
        not_in_warehouse = []
        for day_stock in self.sale_in_day:
            # if the barcode is not exist in the warehouse so add this stock to list
            if not day_stock.is_in_warehouse(self.bar_code_list):
                not_in_warehouse.append(day_stock)
        return not_in_warehouse

    def export_text_file_non_define_stock_in_warehouse(self,not_in_warehouse):
        x = datetime.datetime.now()
        date = x.strftime("%Y-%m-%d-%Hh%M")

        try:
            with open("Output/non_define 2 " + date + ".txt", "w",encoding="utf8") as text:
                for stock in not_in_warehouse:
                    text.write(stock.bar_code + "\n")

            with open("Output/non_define 1 " + date + ".txt", "w",encoding="utf8") as text:
                text.write("\nIn and Out \n")
                for stock in not_in_warehouse:
                    if stock.in_quantity > 0 and stock.sale_quantity > 0:
                        text.write(stock.bar_code + "\n")
                text.write("\nOut \n")
                # write out stock
                for stock in not_in_warehouse:
                    if stock.sale_quantity > 0:
                        text.write(stock.bar_code + "\n")
                # write in stock
                text.write("\nIn \n")
                for stock in not_in_warehouse:
                    if stock.in_quantity > 0:
                        text.write(stock.bar_code + "\n")
        except Exception as e:
            messagebox.showerror("Error", "Can not export file" + str(e))


    def get_sum_in(self):
        sum_in = 0
        for temp_stock in self:
            sum_in += temp_stock.in_stock
        return sum_in

    def get_sum_sale(self):
        sum_sale = 0
        for temp_stock in self:
            sum_sale += temp_stock.sale_stock
        return sum_sale

    def get_sum_ava(self):
        sum_new_ava = 0
        for temp_stock in self:
            sum_new_ava += temp_stock.new_ava_stock
        return sum_new_ava

    def to_data_frame(self):
        dataframe = pandas.DataFrame.from_records([s.to_dict() for s in self])
        print(dataframe)
        return dataframe

    # def visualization(self,is_visualization):
    #     if is_visualization:
    #         dataframe = self.to_data_frame()

    #         watehouse = dataframe.sort_values(by=['new_ava_stock'], ascending=False)
    #         watehouse = watehouse[:50]
    #         print(watehouse)

    #         watehouse.plot(kind='bar', x='vn_name', y='new_ava_stock',figsize=(20,8),title="Warehouse")
    #         plt.xticks(rotation=40)
    #         plt.savefig("Output/Plot/warehouse.jpg")

    #         best_sell = dataframe.sort_values(by=['sale_stock'], ascending=False)
    #         best_sell = best_sell[:50]
    #         best_sell.plot(kind='bar', x='vn_name', y='sale_stock',figsize=(20,8),title="Best seller")

    #         plt.xticks(rotation=40)
    #         plt.savefig("Output/Plot/best_seller.jpg")

    #         in_stock = dataframe.sort_values(by=['in_stock'], ascending=False)
    #         in_stock = in_stock[:50]
    #         in_stock.plot(kind='bar', x='vn_name', y='in_stock',figsize=(20,8),title="Return")

    #         plt.xticks(rotation=40)
    #         plt.savefig("Output/Plot/return.jpg")
    #         plt.show()

