# summary: main.py gets dim log and outputs an Excel table formatted for analysis
 
# import modules
import sys, os, csv
from sys import exit
from os.path import exists
import pandas as pd
from tkinter import Tk, filedialog
from openpyxl import Workbook 
from openpyxl.worksheet.cell_range import CellRange 
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows

class Log:

    def __init__(self):
        # create empty data list
        self.data = []

        # get headers from headers.txt
        self.headers = self.getHeaders()

    def getData(self):
        # replace default icon to log.ico 
        win = Tk()
        win.withdraw()
        win.iconbitmap(r'src/log.ico')

        # open dialog to browse file in File Explorer
        filepath = filedialog.askopenfilename(initialdir='/', 
                                              title="Select a File", 
                                              filetypes=[("Log files", ".txt .log")])
        
        # check if user imported file 
        if filepath != '':
            # append file contents into data list 
            with open(filepath) as logFile:
                for row in csv.reader(logFile, delimiter=';'):
                    self.data.append(row)
        else:
            # error if no file was imported
            print("No log file was imported")
            exit()

    def parseData(self):
        # create data frame and parse through data
        self.getData()
        self.data = pd.DataFrame(self.data).iloc[:,1:-1].dropna()


    def getHeaders(self, line_number=None):
        # specify header file
        filename = 'src/headers.txt'

        # open header file
        with open(filename) as file:
            # get all headers
            if line_number is None:
                return [x.strip(' ') for x in (" ".join(line.strip() for line in file)).split(';')]
            else:
                for i, line in enumerate(file):
                    # get only the headers from the line
                    if i == line_number:
                        return list(filter(None,(line.strip()).split(';')))

class Category(Log):

    def __init__(self, index=None):
        # calls original __init__()
        super().__init__()
        self.headers = self.getHeaders(index)

    def __call__(self,other):
        # get initial headers and its range
        initial_headers = self.getHeaders(0)
        initial_range = [x + 1 for x, header in enumerate(initial_headers)]
        
        # split data to its respective category (i.e. Dimensions, Contour Verify, Corner)
        self.range = initial_range + [index for index, item in enumerate(other.headers) if item in self.headers]
        self.data = (other.data[self.range]).set_axis((initial_headers + self.headers),axis=1)
    
    def formatTable(self, worksheet):
        # define range
        full_range = CellRange(min_col=worksheet.min_column,
                                          min_row=worksheet.min_row,
                                          max_col=worksheet.max_column,
                                          max_row=worksheet.max_row
                                          ).coord

        # set table format
        mediumStyle =TableStyleInfo(name='TableStyleMedium1',
                                    showRowStripes=True)
        # create a table
        table = Table(ref=full_range,
                    displayName=(worksheet.title).replace(" ", "_"),
                    tableStyleInfo=mediumStyle)

        # add the table to the worksheet
        worksheet.add_table(table)

    def paste2Excel(self,worksheet,title):
        # title the worksheet
        worksheet.title = title

        # paste data frame into Excel worksheet
        for rows in dataframe_to_rows(self.data, index=False, header=True):
            worksheet.append(rows)

        # format table from data
        self.formatTable(worksheet)

# add virtual environment to path
sys.path.insert(0, "src/.venv")

# specify excel file name
excel_file = 'dims.xlsx'

# remove if "dims.xlsx" exists
if exists(excel_file):
    os.remove(excel_file)

# initialize log
log = Log()

# parse data from log
log.parseData()

# initialize "Dimensions" category
dim = Category(1)
dim(log)

# initialize "Contour Verify" category
c_v = Category(2)
c_v(log)

# initialize "Corner" category
cor = Category(3)
cor(log)

# open a workbook
wb = Workbook()

# create "Dimensions" sheet & paste data
ws1 = wb.active
dim.paste2Excel(ws1,"Dimensions")


# create "Contour Verify" sheet & paste data
ws2 = wb.create_sheet()
c_v.paste2Excel(ws2,"Contour Verify")

# create "Corner" sheet & paste data
ws3 = wb.create_sheet()
cor.paste2Excel(ws3,"Corner")

# save workbook
wb.save(excel_file)

# check if "dims.xlsx" was created from log
if exists(excel_file):
    print("Log file successfully parsed to Excel!")
else:
    print("ERROR: Log file unsuccessfully parsed to Excel...")
