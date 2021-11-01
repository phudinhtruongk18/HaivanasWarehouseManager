import openpyxl as excel


class ProductOnHand():
    def __init__(self, *args):
        self.bar_code = args[0]
        self.en_name = args[1]
        self.vn_name = args[2]
        self.in_stock = args[3]
        self.sale_stock = args[4]
        self.ava_stock = args[5]
        if not isinstance(self.ava_stock, int):
            print('not int', self.ava_stock)

    def __str__(self):
        return f"\n {self.bar_code} en_name: {self.en_name} + vn_name {self.vn_name} " \
               f"in_stock {self.in_stock}sale_stock {self.sale_stock} ava_stock {self.ava_stock}" + f"\n {self.bar_code} en_name: {self.en_name} + vn_name {self.vn_name} " \
                                                                                                    f"in_stock {type(self.in_stock)}sale_stock {type(self.sale_stock)} ava_stock {type(self.ava_stock)}"


class WareHouse(list):
    def __init__(self, products_on_hand):
        super().__init__()
        for product in products_on_hand:
            self.append(ProductOnHand(*product))

    def show_stock(self):
        for temp in self:
            print(temp)

    def read_data(self):
        print("READ")

    def export(self):
        print("EXPORT")

    def sum_stock(self):
        print("SUM")

    def visualization(self):
        print("VIS")


class ProductInDay():
    def __init__(self, bar_code):
        self.bar_code = bar_code
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

        for temp in self.products_in_day:
            print(temp.bar_code,temp.sale_quantity,temp.in_quantity)
