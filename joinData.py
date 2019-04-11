import openpyxl
from myxlutils import format_date_rows, get_column_names_and_index
# from ForecastManipulations import forecastDict
# from InvoicedManipulations import invoicedDict
# from CreditsManipulations import creditsDict

wbCombined = openpyxl.Workbook()
sCombined = wbCombined.active
sCombined.title = "Detail"

forecast = "Forecast.xlsx"
invoiced = "Invoiced.xlsx"
credit = "Credits.xlsx"

wbFore = openpyxl.load_workbook(forecast)
sFore = wbFore.active

wbInvo = openpyxl.load_workbook(invoiced)
sInvo = wbInvo.active

wbCred = openpyxl.load_workbook(credit)
sCred = wbCred.active

# We will use the column names from Forecast... is this right???
rowCount = 1
for r in range(rowCount, sFore.max_row+1):
        for c in range(1, sFore.max_column+1):
                sCombined.cell(row=r,column=c).value = sFore.cell(row=r, column=c).value
        rowCount+=1
for r in range(rowCount, sInvo.max_row+1):
        for c in range(1, sInvo.max_column+1):
                sCombined.cell(row=r,column=c).value = sInvo.cell(row=r, column=c).value
        rowCount+=1

# NOT WORKING!

for r in range(rowCount, sCred.max_row+1):
        for c in range(1, sCred.max_column+1):
                sCombined.cell(row=r,column=c).value = sCred.cell(row=r, column=c).value
        rowCount+=1

print("Total Rows: " + str(rowCount))
wbCombined.save("Combined.xlsx")
wbCombined.close()
wbCombined = openpyxl.load_workbook("Combined.xlsx")
sCombined = wbCombined.active

combinedDict = {}
get_column_names_and_index(sCombined,combinedDict)
format_date_rows(sCombined, combinedDict, "mm-dd-yy", "Due Date", "InvoiceDateSent", "OriginalDueDate")

wbCombined.save("Combined.xlsx")
wbCombined.close()