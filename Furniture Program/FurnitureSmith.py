import os, openpyxl

os.chdir('c:\\Path')  # File location

wb=openpyxl.load_workbook('Furniture.xlsx')  # Opens the workbook
sheet = wb['Furniture']   # goes to the correct sheet

price_updates = {'E':84.00,'D':184.00,'A':80.40}  #updates the prices

for rowNum in range(4,sheet. max_row+1):       #for statement will go to specific sheet, starting at row 2 until the end of the sheet
    tableType = sheet.cell(row=rowNum, column=5).value   #Declares column value
    if tableType in price_updates:
        sheet.cell(row=rowNum, column=6).value = price_updates[tableType] # if statement explaining if the variable is in the update, it needs to be updated in the workbook

wb.save('updatedFurniture.xlsx')   #Saves the new updated workbook fiile
