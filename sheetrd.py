import os
import csv, xlrd, openpyxl
import collections

import pprint
pp = pprint.PrettyPrinter(indent=4)

'''
why use tuple for rows?
http://stackoverflow.com/questions/626759/whats-the-difference-between-list-and-tuples-in-python

row are heterogeneous, right? but a row, when its with other rows is homogeneous
so... 
rows are tuples (sometimes named tuples)
worksheets are named tuples containing a name and data (data is a list of tuples)
workbooks are lists containing worksheets
'''

class SheetReader(object):

    def __init__(self, file_location, header_row=0):

        try:
            file_handle = open(file_location, 'rb')
        except IOError:
            print "can't open file"


        file_basename = os.path.basename(file_location)
        file_name, file_ext = os.path.splitext(file_basename)

        '''
        using a named tuple for worksheets to contain the sheet's 
        name and data in a simple object
        http://docs.python.org/library/collections.html#collections.namedtuple
        '''
        Worksheet = collections.namedtuple('Worksheet', 
                                           ['sheet_name','sheet_data'])

        if file_ext == '.csv':

            csv_reader = csv.reader(file_handle)

            '''
            TODO
            #what about really big sheets?
            file_size = os.path.getsize(file_location)

            if file_size > some_limit:
                #then what? can't access csv by row/col indexs
            '''

            working_sheet = [] #holds the row tuples
            Drow = None #so that named tuple can be used after its set
        
            for row_num, row in enumerate(csv_reader):
                if header_row is not 0: #args default to 0
                    if row_num == header_row - 1: #enumerate starts at 0

                        header = []

                        '''
                        Sometimes csv's end up with left over columns.
                        Instead of deleteing them here, if there are
                        columns in the header row with no string name
                        they are given a generic name 'col' + index in the list

                        TODO
                        This will need to be expanded to deal with headings
                        that start with numbers or contain problematic chars
                        like $ or %.
                        '''

                        for p_header_num, p_header_cell in enumerate(row):
                            if p_header_cell == '':
                                header.append('col' + str(p_header_num))
                            else:
                                header.append(p_header_cell)
                        
                        #create the named tuple object
                        Drow = collections.namedtuple('Row',header)
                    else:
                        '''
                        not sure why the ._make() works, but it does
                        found example here http://www.daniweb.com/software-development/python/code/286925/exploring-named-tuples-python
                        '''
                        named_row = Drow._make(row)
                        working_sheet.append(named_row)
                else:
                    '''
                    with no header expected a basic tuple of the row
                    is appended to the list of rows
                    ''' 
                    working_sheet.append(tuple(row))

            '''
            the workbook and worksheet construction is used here,
            even though csv's won't have more than one worksheet, to maintain
            consistency in functions that will access the sheet data later
            '''
            workbook = [] #for non-csv, a list of worksheets, here it's only 1
            
            sheet = Worksheet(sheet_name = file_name, 
                                sheet_data = working_sheet)
            
            workbook.append(sheet)
            
            self.sheets = workbook
        
        elif file_ext == '.xls':
            xl_reader = xlrd.open_workbook(file_location)
            
            xl_sheet_names = xl_reader.sheet_names()

            workbook = []

            for xl_sheet_name in xl_sheet_names:

                working_sheet = []
                
                xl_sheet = xl_reader.sheet_by_name(xl_sheet_name)

                for xl_row in range(xl_sheet.nrows):

                    working_sheet.append(xl_sheet.row_values(xl_row))

                sheet = Worksheet(sheet_name = xl_sheet_name, 
                                    sheet_data = working_sheet)

                workbook.append(sheet)

            self.sheets = workbook
        
        elif file_ext == '.xlsx':
            xlx_reader = openpyxl.load_workbook(filename = file_location, 
                                                use_iterators=True)

            xlx_sheet_names = xlx_reader.get_sheet_names()

            workbook = []

            for xlx_sheet_name in xlx_sheet_names:

                working_sheet = []

                xlx_sheet = xlx_reader.get_sheet_by_name(xlx_sheet_name)

                for xlx_row in xlx_sheet.iter_rows():

                    row_values = []

                    for cell in xlx_row:
                        row_values.append(cell.internal_value)

                    working_sheet.append(row_values)

                sheet = Worksheet(sheet_name = xlx_sheet_name, 
                                    sheet_data = working_sheet)

                workbook.append(sheet)

            self.sheets = workbook
        
        else:
            print "unsupported file type. Only .csv, .xls and .xlsx"


    def book_names(self):
        for sheet in self.sheets:
            yield sheet.sheet_name

    def sheet_rows(self, sheet_name):
        for sheet in self.sheets:
            if sheet.sheet_name == sheet_name:
                for row in sheet.sheet_data:
                    yield row

    def col(self, sheet, col_choice):
        pass