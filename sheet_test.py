#!/usr/bin/env python

import sheetrd

import pprint
pp = pprint.PrettyPrinter(indent=4)


csv_wb = sheetrd.SheetReader('test_data/csv_short.csv', header=[True],header_row=[1])

#xls_wb = sheetrd.SheetReader('test_data/xls_testing.xls',header=[True,True], header_row=[1,1])
xls_wb = sheetrd.SheetReader('test_data/xls_testing.xls')

xlsx_wb = sheetrd.SheetReader('test_data/xlsx_testing.xlsx')
#xlsx_wb = sheetrd.SheetReader('test_data/xlsx_testing.xlsx',header=[True,True,False], header_row=[1,1,None])


#pp.pprint(csv_wb.sheets)
#pp.pprint(xls_wb.sheets)
#pp.pprint(xlsx_wb.sheets)


#pp.pprint(csv_wb.book_names())
#pp.pprint(xls_wb.book_names())
#pp.pprint(xlsx_wb.book_names())


#for row in csv_wb.sheet_rows():
#	print row

#for row in xls_wb.sheet_rows('testing'):
#	print row

#for row in xls_wb.sheet_rows():
#    print row

#for row in xlsx_wb.sheet_rows('testing_xlsx','sheetEnd'):
#    print row

#for row in xlsx_wb.sheet_rows():
#    print row



#for col in xlsx_wb.sheet_cols('testing_xlsx'):
#   print col

#for col in xlsx_wb.sheet_cols('testing_xlsx','sheetEnd'):
#   print col

xlsx_df = xlsx_wb.convert_to_dataframe('testing_xlsx')
pp.pprint(xlsx_df)
