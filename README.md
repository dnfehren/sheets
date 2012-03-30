#Sheets
a python object for seamless spreadsheet reading across formats

##Goals
Allow a python application to open and read data from a spreadsheet file without the need for specific file type coding. Users should be able to work with data from .csv, Excel .xls or .xlsx without needing to specifically create or invoke the logic from file specific libraries. Contributions welcome.

##How it works
Spreadsheet files are identified based on their extension and the appropriate module is used to read the data from the file into a python data structure.

* csv is used for .csv files
* xlrd is used for .xls files
* openpyxl is used for .xlsx files

The sheets object that is created looks like this...

    +-------------------------------------------------------+
    |Workbook = list                                        |
    |-------------------------------------------------------|
    |                                                       |
    |  +--------------------------------------------------+ |
    |  |Worksheet = named tuple                           | |
    |  |--------------------------------------------------| |
    |  |                                                  | |
    |  |  +---------------------------------------------+ | |
    |  |  |sheet_name = string                          | | |
    |  |  +---------------------------------------------+ | |
    |  |                                                  | |
    |  |  +---------------------------------------------+ | |
    |  |  |sheet_data = list                            | | |
    |  |  |---------------------------------------------| | |
    |  |  |                                             | | |
    |  |  |  +----------------------------------------+ | | |
    |  |  |  |Row =  tuple or named tuple             | | | |
    |  |  |  |----------------------------------------| | | |
    |  |  |  |                                        | | | |
    |  |  |  |                                        | | | |
    |  |  |  +----------------------------------------+ | | |
    |  |  +---------------------------------------------+ | |
    |  +--------------------------------------------------+ |
    +-------------------------------------------------------+


##Examples
See sheet_test.py for some (very basic) working examples.

##Future Plans
*See issues

##Author
Dan Fehrenbach
dnfehrenbach@gmail.com
