import os
import tkinter as jra
from tkinter import messagebox, ttk
from tkinter.filedialog import askopenfilename
from tkinter import font as tkfont

from ExcelManager import read_warehouse
from ObjCanThiet import WareHouse, SaleData
from helper import create_export_fordel


def openThuMuc(link=""):
    os.startfile(f'{os.path.realpath("") + link}')


class Application(jra.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        master.title("WareHouseManager")
        master.iconbitmap("logo.ico")
        w = 420
        h = 500
        ws = master.winfo_screenwidth()
        hs = master.winfo_screenheight()
        x = ws / 2 - (w / 2)
        y = (hs / 2) - (h / 2)
        master.geometry('%dx%d+%d+%d' % (w, h, x, y))

        self.canvas_batdau = jra.Frame(self)

        self.font = tkfont.Font(family='Helvetica', size=12, weight="bold")

        jra.Label(self.canvas_batdau, text="").grid(row=1, column=0)

        self.pick_warehouse = jra.Button(self.canvas_batdau, text="Pick Warehouse", font=self.font, fg="#ffffff",
                                          bg="#263942",
                                          width=20,
                                          height=2)
        self.pick_warehouse["command"] = self.pick_warehouse_file
        self.pick_warehouse.grid(row=2, column=0)

        jra.Label(self.canvas_batdau, text="").grid(row=3, column=0)

        self.pick_day_data = jra.Button(self.canvas_batdau, text="Pick Day Sale", font=self.font, fg="#ffffff",
                                        bg="#263942",
                                        width=20,
                                        height=2)
        self.pick_day_data["command"] = self.pick_day_file
        self.pick_day_data.grid(row=4, column=0)

        jra.Label(self.canvas_batdau).grid(row=5, column=0)

        self.string_selected_sheet = jra.StringVar(self.master, value="Select Sheet")
        self.sheet_option_menu = jra.OptionMenu(self.canvas_batdau, self.string_selected_sheet, None)
        self.sheet_option_menu.configure(font=self.font, width=18, anchor=jra.CENTER, bd=5)
        self.menu_option = self.master.nametowidget(self.sheet_option_menu.menuname)
        self.menu_option.config(font=self.font)

        self.sheet_option_menu.grid(row=7, column=0)

        jra.Label(self.canvas_batdau).grid(row=8, column=0)

        self.is_visualization = jra.BooleanVar(self)
        self.check_box = jra.Checkbutton(self, text="Visualization", variable=self.is_visualization)
        self.check_box.grid(row=0, column=0)

        self.buttonSoSanh = jra.Button(self.canvas_batdau, text="Thực thi", font=self.font, fg="#ffffff",
                                       bg="#263942",
                                       width=20,
                                       height=2)
        self.buttonSoSanh["command"] = self.kiem_tra_va_thuc_thi
        self.buttonSoSanh.grid(row=10, column=0)

        jra.Label(self.canvas_batdau).grid(row=11, column=0)

        self.buttonThuMuc = jra.Button(self.canvas_batdau, text="Mở Thư Mục Hiện Hành ", font=self.font, fg="#ffffff",
                                       bg="#263942",
                                       width=20, height=2)
        self.buttonThuMuc["command"] = openThuMuc
        self.buttonThuMuc.grid(row=12, column=0)

        self.label_trang_thai = jra.Label(self.canvas_batdau, text="Downloading and uploading your files")

        self.progress_bar = ttk.Progressbar(self.canvas_batdau, orient=jra.HORIZONTAL, length=100, mode="determinate")

        self.huongDan = jra.Label(self.canvas_batdau, text="\nBất kì vấn đề nào \n phudinhtruongk18@gmail.com", width=60)
        self.huongDan.grid(row=13, column=0)

        self.canvas_batdau.grid()

        self.warehouse_file = None
        self.day_sale_file = None
        self.stock_in_day = None
        self.linkFile = os.path.realpath("")

        self.grid()
        self.mainloop()

    def pick_day_file(self):
        fileName = askopenfilename(defaultextension='.xlsx', initialdir=self.linkFile)
        temp_name = str(fileName).split("/")[-1]
        self.pick_day_data.configure(text=temp_name)
        self.day_sale_file = fileName

        self.stock_in_day = SaleData(fileName)
        self.reset_saved_sessions(self.stock_in_day.sheet_names)

    def pick_warehouse_file(self):
        fileName = askopenfilename(defaultextension='.xlsx', initialdir=self.linkFile)
        temp_name = str(fileName).split("/")[-1]
        self.pick_warehouse.configure(text=temp_name)
        self.warehouse_file = fileName

    def reset_saved_sessions(self,session_list):
        # change database here
        self.menu_option.delete(0, "end")
        for session in session_list:
            # set session.ID for trace after select in option menu
            self.menu_option.add_command(label=session, command=lambda value=session: self.string_selected_sheet.set(value))

    def kiem_tra_va_thuc_thi(self):

        create_export_fordel("./Output")
        create_export_fordel("./Output/Plot")

        self.stock_in_day = SaleData(self.day_sale_file)

        print(self.string_selected_sheet)
        self.stock_in_day.get_list_product(self.string_selected_sheet.get())

        print(self.stock_in_day.products_in_day.__len__())

        raw_data = read_warehouse(self.warehouse_file, "SOH GOOD")
        warehouse = WareHouse(products_on_hand=raw_data)

        # they are romeo and juliet in 2021
        warehouse.set_sale_in_day(self.stock_in_day.products_in_day)
        # check stock not in warehouse
        stock_non_def = warehouse.get_stocks_not_in_warehouse()
        print(stock_non_def)
        if stock_non_def.__len__() > 0:
            messagebox.showinfo("Check your warehouse", "There are non define bar code!")
            warehouse.export_text_file_non_define_stock_in_warehouse(stock_non_def)
            return

        warehouse.cuculate_stock()
        summ_in = warehouse.get_sum_in()
        summ_sale = warehouse.get_sum_sale()
        summ_ava = warehouse.get_sum_ava()
        warehouse.export(summ_in, summ_sale, summ_ava)

        warehouse.visualization(self.is_visualization.get())

        messagebox.showinfo("Done","Complete! Thank for using aram tool")


print("Hello")
giaoDien = jra.Tk()

app = Application(master=giaoDien)
