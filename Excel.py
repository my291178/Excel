
# rb = xlrd.open_workbook('../ArticleScripts/ExcelPython/xl.xls', formatting_info=True)
#
# sheet = rb.sheet_by_index(0)



import xlrd, xlwt
import openpyxl

from openpyxl import Workbook, load_workbook
from openpyxl.styles.borders import Border, Side

thick_border = Border(left=Side(style='thick'),
                      right=Side(style='thick'),
                      top=Side(style='thick'),
                      bottom=Side(style='thick'))


wb = load_workbook('report.xlsx')
ws = wb.active
ws.title = 'test'

ws['A1'] = "ОТЧЕТ"
ws.merge_cells('A1:D1')
#ws.add_image()

ws.min_column

ws['A1'].border = thick_border
ws['B1'].border = thick_border
ws['C1'].border = thick_border
ws['D1'].border = thick_border

ws['A2'] = "ID"
ws['A2'].border = thick_border

ws['B2'] = "Name"
ws['B2'].border = thick_border

ws['C2'] = "Date"
ws['C2'].border = thick_border

ws['D2'] = "Price"
ws['D2'].border = thick_border

wb.save('report.xlsx')
