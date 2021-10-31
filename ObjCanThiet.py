
class ProductOnHand():
    def __init__(self, *args):
        self.bar_code = args[0]
        self.en_name = args[1]
        self.vn_name = args[2]
        self.in_stock = args[3]
        self.sale_stock = args[4]
        self.ava_stock = args[5]
        if not isinstance(self.ava_stock, int):
            print('not int',self.ava_stock)

    def __str__(self):
        return f"\n {self.bar_code} en_name: {self.en_name} + vn_name {self.vn_name} " \
               f"in_stock {self.in_stock}sale_stock {self.sale_stock} ava_stock {self.ava_stock}"+f"\n {self.bar_code} en_name: {self.en_name} + vn_name {self.vn_name} " \
               f"in_stock {type(self.in_stock)}sale_stock {type(self.sale_stock)} ava_stock {type(self.ava_stock)}"


class ProductInDay():
    def __init__(self, *args):
        self.bar_code = args[0]
        self.is_sell = args[1]
        self.quantity = args[2]


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

