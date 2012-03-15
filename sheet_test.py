import sheetrd

sheet = sheetrd.SheetReader('testing.xlsx')

print sheet.sheets[0]['data'][4]