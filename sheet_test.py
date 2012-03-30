import sheetrd

import pprint
pp = pprint.PrettyPrinter(indent=4)


csv_wb = sheetrd.SheetReader('test_data/csv_testing.csv')
xls_wb = sheetrd.SheetReader('test_data/xls_testing.xls', 1)
xlsx_wb = sheetrd.SheetReader('test_data/xlsx_testing.xlsx', 1)

#pp.pprint(xls_wb.sheets)

for row in csv_wb.sheet_rows('csv_testing'):
	print row

#for row in xls_wb.sheet_rows('testing'):
#	print row

#for name in xlsx_wb.book_names():
#	for row in xlsx_wb.sheet_rows(name):
#		print row