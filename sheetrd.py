import os
import csv, xlrd, openpyxl

class SheetReader(object):

    def __init__(self, file_location):

        try:
            file_handle = open(file_location, 'rb')
        except IOError:
            print "can't open file"

        file_path_and_name, file_ext = os.path.splitext(file_location)

        if file_ext == '.csv':

            csv_reader = csv.reader(file_handle)

            '''
            #what about really big sheets?
            file_size = os.path.getsize(file_location)

            if file_size > some_limit:
                #then what? can't access csv by row/col indexs
            '''

            working_sheet = []
        
            for row in csv_reader:
                working_sheet.append(row)

            workbook = []
            sheet = {}
            sheet['name'] = 'sheet1'
            sheet['data'] = working_sheet
            workbook.append(sheet)
            
            self.sheets = workbook
        
        elif file_ext == '.xls':
            xl_reader = xlrd.open_workbook(file_location)
            
            xl_sheet_names = xl_reader.sheet_names()

            workbook = []

            for xl_sheet_name in xl_sheet_names:

                sheet = {}
                sheet['name'] = xl_sheet_name
                sheet['data'] = []
                
                xl_sheet = xl_reader.sheet_by_name(xl_sheet_name)

                for xl_row in range(xl_sheet.nrows):

                    sheet['data'].append(xl_sheet.row_values(xl_row))

                workbook.append(sheet)

            self.sheets = workbook
        
        elif file_ext == '.xlsx':
            xlx_reader = openpyxl.load_workbook(filename = file_location, use_iterators=True)

            xlx_sheet_names = xlx_reader.get_sheet_names()

            workbook = []

            for xlx_sheet_name in xlx_sheet_names:

                sheet = {}
                sheet['name'] = xlx_sheet_name
                sheet['data'] = []
                
                xlx_sheet = xlx_reader.get_sheet_by_name(xlx_sheet_name)

                for xlx_row in xlx_sheet.iter_rows():

                    row_values = []

                    for cell in xlx_row:
                        row_values.append(cell.internal_value)

                    sheet['data'].append(row_values)

                workbook.append(sheet)

            self.sheets = workbook
        
        else:
            print "unsupported file type .csv, .xls and .xlsx"


    def row(self, sheet, row_choice):
        self.sheets


    def col(self, sheet, col_choice):
        pass