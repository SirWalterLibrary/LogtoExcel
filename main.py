# summary: main.py gets dim log and outputs an Excel table formatted for analysis
 
# import modules
import os
from os.path import exists
from functions import Data
from openpyxl import Workbook 

excel_file = 'dims.xlsx'

# remove if "dims.xlsx" exists
if exists(excel_file):
    os.remove(excel_file)

# initialize log
log = Data()

# parse data from log
log.parseData()

# initialize "Dimensions" category
dim = Data(1)
dim(log)

# initialize "Contour Verify" category
c_v = Data(2)
c_v(log)

# initialize "Corner" category
cor = Data(3)
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
