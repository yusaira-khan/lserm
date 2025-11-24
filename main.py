import glob

import openpyxl

from input_sheet import InputSheet
from output_excel import OutputExcel
from output_gsheet import OutputGsheet

WINDOWS_GLOB = "C:\\Users\\yusai\\Downloads\\Q*.xlsx"
MAC_GLOB = "/Users/yusairak/Downloads//Q*.xlsx"

if __name__ == "__main__":
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for file in glob.glob(MAC_GLOB):
        print("using file", file)
        i = InputSheet(file).load()
        # o = OutputExcel.convert(i.output, wb)
        o = OutputGsheet.convert(i.output, wb)
        o.write()
    # wb.save("lserm.xlsx")
    # wb.close()
