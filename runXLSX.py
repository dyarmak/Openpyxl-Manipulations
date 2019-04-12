# 1. read all the files in the directory
# 1a. convert any .xls files to .xlsx
# 2. load each .xlsx file into a variable
# 3. Perform Christy's manipulations
# Amalgamate the 3 queries into one?
# 4. Open the old ForecastYYYYMMDD.xlsx file
# 4a. rename to todays date
# 5. Delete the data in the detail sheet 
# 6. copy - paste contents from new queries  

import os
import shutil
from os import path

# FileName variables
forecastFName = "qry_Forecast.xlsx"
invoicedFName = "qry_Invoiced.xlsx"
creditFName = "qry_Credits.xlsx"
combinedFName = "Combined.xlsx"

# ******************* PLEASE NOTE **************************
# OpenPyXL can only work with .xlsx files (Excel 2007 and newer)
# IF the files are of .xls type, execute the runXLS.py script
# It will convert from .xls to .xlsx

print("*** WARNING ***")
print("This script will look in the current folder for\nqry_Forecast_cc.xlsx\nqry_CreditsCC.xlsx\nqry_invoicedcc.xlsx")
print("If these files are not present, or are named differently it may not work properly!")
input("If files are good, Press Enter to continue")

# Make Output Folder
savePath = "py_Output"
if os.path.exists(savePath) is False:
        os.mkdir(savePath)


if os.path.exists("qry_Forecast_cc.xlsx"):
    src = os.path.realpath("qry_Forecast_cc.xlsx")
    head, tail = path.split(src)
    # print("Path: " + head)
    # print("File: " + tail)
    dstFolder = head + "\\" + savePath
    # print("Dst folder: " + dstFolder)
    tail = forecastFName
    dst = dstFolder + "\\" + tail
    shutil.copy(src, dst)
else:
    print("qry_forecast_cc.xlsx does not exits in this directory")
    print("Check file extension.\n IF it is .xls, run the other module")

if os.path.exists("qry_CreditsCC.xlsx"):
    src = os.path.realpath("qry_CreditsCC.xlsx")
    head, tail = path.split(src)
    # print("Path: " + head)
    # print("File: " + tail)
    dstFolder = head + "\\" + savePath
    # print("Dst folder: " + dstFolder)
    tail = creditFName
    dst = dstFolder + "\\" + tail
    shutil.copy(src, dst)
else:
    print("qry_CreditsCC.xlsx does not exits in this directory")
    print("Check file extension.\n IF it is .xls, run the other module first")

if os.path.exists("qry_invoicedcc.xlsx"):
    src = os.path.realpath("qry_invoicedcc.xlsx")
    head, tail = path.split(src)
    # print("Path: " + head)
    # print("File: " + tail)
    dstFolder = head + "\\" + savePath
    # print("Dst folder: " + dstFolder)
    tail = invoicedFName
    dst = dstFolder + "\\" + tail
    shutil.copy(src, dst)
else:
    print("qry_forecast_cc.xlsx does not exits in this directory")
    print("Check file extension.\n IF it is .xls, run the other module")

# Change to the py_output folder for the remaining operations 
os.chdir(savePath)

# RUN Forecast manipulation code
import ForecastManipulations
# RUN Credits manipulation code
import CreditsManipulations
# RUN Invoiced manipulation code
import InvoicedManipulations
# Amalgamate the three files into one
# import joinData
# Open .xlsx file with pivots

# Clear contents of details sheet

# Copy the alalgamation into the details tab

# Add VLookup values

# Fix WHERE Type == HSE || CRB || Software 