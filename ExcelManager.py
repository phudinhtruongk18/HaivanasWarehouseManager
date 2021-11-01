import datetime
import openpyxl as excel
import os
from openpyxl.styles import Font, Alignment, GradientFill
from openpyxl.utils import get_column_letter


def read_warehouse(fileName,sheetname):
    # time a lot of time
    docData = excel.load_workbook(filename=fileName)
    duLieu = docData[sheetname]
    listDuLieuTemp = []
    for index, dataTemp in enumerate(duLieu.values):
        # bo ba hang đau tien
        if index < 3:
            continue
        # ket thuc neu xuat hien chu nay
        if dataTemp[0] is None or dataTemp[0] == "END_WAREHOUSE":
            break
        # lay du lieu tu 7 cot
        listDuLieuTemp.append(dataTemp[:7])
    return listDuLieuTemp


def XuatFileExcel(list_xuat):
    ghiData = excel.Workbook()
    trangTinh = ghiData.active
    trangTinh.title = "Sheet1"

    trangTinh.append(("        URL        ", "Meta", "                                                 Content                                  "))
    a = trangTinh['A1']
    b = trangTinh['B1']
    c = trangTinh['C1']

    a.font = Font(size=16, bold=True)
    b.font = Font(size=16, bold=True)
    c.font = Font(size=16, bold=True)
    ft = Font(size=11)

    # fill_father = GradientFill(stop=("FA4039", "FA4039"))
    # fill_son = GradientFill(stop=("FFAD19", "FFAD19"))

    adjust_column_width_from_col(trangTinh, 1, 1, trangTinh.max_column)
    stt = 2

    mau = True
    for index, object_nek in enumerate(list_xuat):
        # if mau:
        #     fill_chon = fill_father
        # else:
        #     fill_chon = fill_son
        # mau = not mau

        for doituong_matches in object_nek.excel_format():
            print(doituong_matches)
            trangTinh.append((
                doituong_matches
            ))

            a = trangTinh[f'A{stt}']
            b = trangTinh[f'B{stt}']
            c = trangTinh[f'C{stt}']
            stt += 1
            # a.fill = fill_chon
            # b.fill = fill_chon
            # c.fill = fill_chon
            a.font = ft
            b.font = ft
            c.font = ft

            # a.alignment = Alignment(horizontal="center", vertical="center")
            # b.alignment = Alignment(horizontal="center", vertical="center")
            # c.alignment = Alignment(horizontal="center", vertical="center")

    linkthumuc = os.curdir
    x = datetime.datetime.now()
    date = x.strftime("%Y-%m-%d-%Hh%M")
    ghiData.save(filename=linkthumuc + "/Output/" + date + ".xlsx")


def adjust_column_width_from_col(ws, min_row, min_col, max_col):
    column_widths = []

    for i, col in \
            enumerate(
                ws.iter_cols(min_col=min_col, max_col=max_col, min_row=min_row)
            ):

        for cell in col:
            value = cell.value
            if value is not None:

                if isinstance(value, str) is False:
                    value = str(value)

                try:
                    column_widths[i] = max(column_widths[i], len(value))
                except IndexError:
                    column_widths.append(len(value))

    for i, width in enumerate(column_widths):
        col_name = get_column_letter(min_col + i)
        value = column_widths[i] + 10
        ws.column_dimensions[col_name].width = value