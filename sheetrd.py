#!/usr/bin/env python

import os, collections, itertools, re

import pprint
pp = pprint.PrettyPrinter(indent=4)

#Intro

'''
Wrapper class for reading common spreadsheet types. 
CSV, XLS and XLSX are supported currently

Dan Fehrenbach
dnfehrenbach@gmail.com

dependencies are collections (for named tuples), csv, xlrd and openpyxl

why use tuple for rows?
http://stackoverflow.com/questions/626759/whats-the-difference-between-list-and-tuples-in-python

row are heterogeneous, right? but a row, when its with other rows is homogeneous
so... 
rows are tuples (sometimes named tuples)
worksheets are named tuples containing a name and data (data is a list of tuples)
workbooks are lists containing worksheets
'''

#Work
#Functions

def make_header(potential_header_row):
'''
check a row from a spreadsheet and create a list suitable for use as
the elements of a named tuple
'''
    rx_cleaner = re.compile("['(',')','$','-']")
    rx_space = re.compile("\s+")
    rx_first_digit = re.compile("^\d")

    clean_header = []

    for p_head_num, p_head_cell in enumerate(potential_header_row):
        if p_head_cell == '':
            clean_header.append('col' + str(p_head_num))
        else:

            #remove special characters
            clean_head_cell = rx_cleaner.sub('',p_head_cell).strip()
            
            #replace spaces with '_'
            spaced_head_cell = rx_space.sub('_',clean_head_cell).strip()

            final_head_cell = ''

            #if the header starts with a number, append a 'd_'
            # named tuple fields cannot start with a number or '-'
            # but if you use an ordereddict, this might not be necessary
            if rx_first_digit.match(spaced_head_cell):
                final_head_cell = 'd_' + spaced_head_cell
            else:
                final_head_cell = spaced_head_cell

            clean_header.append(final_head_cell)

        return clean_header




'''
take a row from the openpyxl iterator and return just the values

this (likely) depends on the row being generated from the iter_rows()
method that comes with using an openpyxl reader with use_iterators=True
'''
def xlsx_row_values(openpyxl_row):

    values_only_row = []

    for cell in openpyxl_row:
        values_only_row.append(cell.internal_value)

    return values_only_row




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

        '''
        CSV Handling
        '''
        if file_ext == '.csv':

            import csv

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

                        '''
                        Sometimes csv's end up with left over columns.
                        Instead of deleteing them here, if there are
                        columns in the header row with no string name
                        they are given a generic name 'col' + index in the list
                        '''
                        header = make_header(row)
                        
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
            
            import xlrd

            xl_reader = xlrd.open_workbook(file_location)
            
            xl_sheet_names = xl_reader.sheet_names()

            workbook = []

            for xl_sheet_name in xl_sheet_names:

                working_sheet = []
                Drow = None
                
                xl_sheet = xl_reader.sheet_by_name(xl_sheet_name)

                '''
                to get xlrd row values you need an index number,
                AFAIK there isnt a way to iterate by row w/o and index
                '''
                for xl_row_num in range(xl_sheet.nrows):
                    
                    if header_row is not 0:
                    
                        if xl_row_num == header_row - 1:          
                            p_head_row = xl_sheet.row_values(xl_row_num)
                            header = make_header(p_head_row)
                            Drow = collections.namedtuple('Row',header)
                        else:
                            xl_row = xl_sheet.row_values(xl_row_num)
                            named_row = Drow._make(xl_row)
                            working_sheet.append(named_row)
                    else:
                        xl_row = xl_sheet.row_values(xl_row_num)
                        working_sheet.append(tuple(xl_row))

                sheet = Worksheet(sheet_name = xl_sheet_name, 
                                    sheet_data = working_sheet)

                workbook.append(sheet)

            self.sheets = workbook


        elif file_ext == '.xlsx':

            import openpyxl

            xlsx_reader = openpyxl.load_workbook(filename = file_location, 
                                                use_iterators=True)

            xlsx_sheet_names = xlsx_reader.get_sheet_names()

            workbook = []

            for xlsx_sheet_name in xlsx_sheet_names:

                working_sheet = []
                Drow = None

                xlsx_sheet = xlsx_reader.get_sheet_by_name(xlsx_sheet_name)

                '''
                openpyxl doesnt have a row values function, cells need to have
                their values extracted individually, using sheet.iter_rows() 
                should make things as fast as possible
                REF: http://packages.python.org/openpyxl/optimized.html
                '''
                for xlsx_row_num, xlsx_row in enumerate(xlsx_sheet.iter_rows()):

                    if header_row is not 0:

                        if xlsx_row_num == header_row - 1:
                            p_head_row = xlsx_row_values(xlsx_row)
                            header = make_header(p_head_row)
                            Drow = collections.namedtuple('Row',header)
                        else:
                            xlsx_row = xlsx_row_values(xlsx_row)
                            named_row = Drow._make(xlsx_row)
                            working_sheet.append(named_row)
                    else:
                        xlsx_row = xlsx_row_values(xlsx_row)
                        working_sheet.append(tuple(xlsx_row))

                sheet = Worksheet(sheet_name = xlsx_sheet_name, 
                                    sheet_data = working_sheet)

                workbook.append(sheet)

            self.sheets = workbook
        
        else:
            print "unsupported file type. Only .csv, .xls and .xlsx"


    #returns the names of all sheets in a book
    def book_names(self):
        b_names = [x.sheet_name for x in self.sheets]
        return b_names

    #returns a row iterator for the specified sheet
    def sheet_rows(self, sheet_name):
        for sheet in self.sheets:
            if sheet.sheet_name == sheet_name:
                for row in sheet.sheet_data:
                    yield row

    #returns a column iterator for the specified sheet
    #this is based on code in the python docs but I don't get what the * does
    def sheet_cols(self, sheet_name):
        for sheet in self.sheets:
            if sheet.sheet_name == sheet_name:
                for col in itertools.izip_longest(*sheet.sheet_data):
                    yield col
    
    '''
    #http://code.activestate.com/recipes/192401-quickly-remove-or-order-columns-in-a-list-of-lists/
    #http://stackoverflow.com/questions/1983902/remove-row-or-column-from-2d-list-if-all-values-in-that-row-or-column-are-none
    def del_col(self, sheet_name, col_num_list):
        for sheet in self.sheets:
            if sheet.sheet_name == sheet_name:
                pass
                #go through each row and delete something at specific index

                #delete a whole column
    '''

    def convert_to_databook(self):
        
        try:
            import tablib
        except ImportError:
            print "problems importing tablib, install it please"

        tblb_databook = tablib.Databook()

        for sheet in self.sheets:
            pass
            #tblb_dataset = tablib.Dataset(self.sheet_data)
            #tblb_databook
