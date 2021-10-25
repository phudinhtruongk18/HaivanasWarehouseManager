import os
import tkinter as jra
import tkinter.messagebox
from threading import Thread
from tkinter import messagebox, ttk
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter import font as tkfont

from test import thuc_thi_cong_viec


def openThuMuc(link=""):
    os.startfile(f'{os.path.realpath("") + link}')


class Application(jra.Frame):
    def __init__(self, master=None,api_key=None):
        super().__init__(master)
        self.api_key = api_key
        master.title("WareHouseManager")
        master.iconbitmap("logo.ico")
        w = 555
        h = 500
        ws = master.winfo_screenwidth()
        hs = master.winfo_screenheight()
        x = ws / 2 - (w / 2)
        y = (hs / 2) - (h / 2)
        master.geometry('%dx%d+%d+%d' % (w, h, x, y))

        self.canvas_batdau = jra.Frame(self)

        self.labelllll = jra.Label(self.canvas_batdau, text="Num of Threads")
        self.labelllll.grid(row=1, column=0)

        self.font = tkfont.Font(family='Helvetica', size=12, weight="bold")

        self.entry_link = jra.Text(self.canvas_batdau, font=tkfont.Font(family='Helvetica', size=14, weight="bold"),
                                   fg="#000000",
                                   bg="#B0F5AB", width=50, height=10)
        self.entry_link.grid(row=4, column=0)


        jra.Label(self.canvas_batdau, text="Your link").grid(row=3, column=0)

        self.entry_api = jra.Entry(self.canvas_batdau, font=tkfont.Font(family='Helvetica', size=14, weight="bold"),
                                   fg="#000000",
                                   bg="#B0F5AB", width=50)
        self.entry_api.grid(row=2, column=0)

        self.entry_api.insert(1, self.api_key)

        jra.Label(self.canvas_batdau).grid(row=5, column=0)
        self.buttonSoSanh = jra.Button(self.canvas_batdau, text="Thực thi", font=self.font, fg="#ffffff",
                                       bg="#263942",
                                       width=20,
                                       height=2)
        self.buttonSoSanh["command"] = self.kiem_tra_va_thuc_thi
        self.buttonSoSanh.grid(row=6, column=0)

        jra.Label(self.canvas_batdau).grid(row=10, column=0)

        self.buttonThuMuc = jra.Button(self.canvas_batdau, text="Mở Thư Mục Hiện Hành ", font=self.font, fg="#ffffff",
                                       bg="#263942",
                                       width=20, height=2)
        self.buttonThuMuc["command"] = openThuMuc
        self.buttonThuMuc.grid(row=11, column=0)

        self.label_trang_thai = jra.Label(self.canvas_batdau, text="Downloading and uploading your files")

        self.progress_bar = ttk.Progressbar(self.canvas_batdau, orient=jra.HORIZONTAL, length=100, mode="determinate")


        self.huongDan = jra.Label(self.canvas_batdau, text="\nBất kì vấn đề nào \n phudinhtruongk18@gmail.com", width=60)
        self.huongDan.grid(row=12, column=0)

        self.canvas_batdau.grid()

        self.grid()
        self.mainloop()

    def kiem_tra_va_thuc_thi(self):
        if not os.path.exists("Output"):
            os.mkdir("Output")
        try:
            api_key = int(self.entry_api.get())
        except Exception as e:
            messagebox.showinfo("Check threads number ",e)
            return
        if api_key < 0:
            messagebox.showinfo("Check threads number","Check threads number")
            return
        print("Your thread", api_key)
        link_tonghop = self.entry_link.get('1.0', 'end')
        list_link = []
        for temp in str(link_tonghop).split("\n"):
            list_link.append(temp.strip())
        print(list_link)

        thuc_thi_cong_viec(api_key,list_link)
        messagebox.showinfo("Done","Complete! Thank for using my tool")

    def thuc_thi_khi_hoantat(self, ten_file):
        self.canvas_batdau.grid_forget()
        self.canvas_ketuqua.grid()
        self.Text_ketqua.delete('1.0', jra.END)
        self.entry_link.delete('1.0', jra.END)
        fileCanMo = open(ten_file, "r")
        dulieu = fileCanMo.read()
        fileCanMo.close()
        self.Text_ketqua.insert('1.0', dulieu)


print("Hello")
giaoDien = jra.Tk()
apikey = ""

app = Application(master=giaoDien,api_key=str(apikey.strip()))