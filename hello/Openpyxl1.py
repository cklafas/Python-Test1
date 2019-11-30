#import openpyxl
import xlwings as xw
#load workbook
# wk = openpyxl.load_workbook('ExcelGenerated1.xlsx')

# wk = openpyxl.load_workbook('LibreSheet1.xlsx')
# print (wk.sheetnames)

wb = xw.Book('LibreSheet1.xlsx')
sht2 = wb.sheets['Sheet3']
sht3 = wb.sheets['Fruit]']
sht2.range('B2').value = 45

#ss=openpyxl.load_workbook("LibreSheet1.xlsx")
#printing the sheet names

wb.sheets.add('Fruit')


for sheet in wb.sheets:
    if 'Fruit' in sheet.name: 
        sheet.delete()