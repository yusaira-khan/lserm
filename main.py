import glob

import openpyxl

from input_sheet import InputSheet
from lserm.output_excel import OutputExcel

UNUSED="""
LSERM DOC
| col  |    A   |   B   |  C    |
|------+--------+-------+-------|
| pt   | 365.45 | 98.55 | 48.15 |
| pica |  30.37 |  8.21 |  4.01 |
| code |  65.67 | 17.76 | 13.00 |
"""
if __name__ == "__main__":
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for file in glob.glob("C:\\Users\\yusai\\Downloads\\Q*.xlsx"):
        print("using file",file)
        i = InputSheet(file)
        i.load()
        i.output.__class__ = OutputExcel
        i.output.write(wb)
    wb.save("lserm.xlsx")
    wb.close()

