"""Rebuild the synthetic BIFF8 fixture with xlwt==1.3.0 (test generation only)."""
from pathlib import Path
import xlwt

book = xlwt.Workbook()
sheet = book.add_sheet('Markets')
sheet.write_merge(0, 0, 0, 2, 'Inventory')
for column, value in enumerate(['SKU', 'Description', 'Units']):
    sheet.write(1, column, value)
sheet.write(2, 0, 'JO-1')
sheet.write(2, 2, 42)
sheet.write(3, 0, 'EG-2')
sheet.write(3, 1, 'مرحبا 한국어')
sheet.write(3, 2, 55)
sheet.write(4, 2, xlwt.Formula('C3+C4'))
other = book.add_sheet('Secondary')
other.write(0, 0, 'IL')
other.write(0, 1, True)
book.save(str(Path(__file__).with_name('structured.xls')))
