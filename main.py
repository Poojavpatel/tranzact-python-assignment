import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

path = 'data.xlsx'
wb=load_workbook(path)

level = {
  ".1" : {'row':0, 'column':-1},
  "..2" : {'row':-1, 'column':1},
  "..3" : {'row':-1, 'column':1}
}

# Function to create a sheet
def createSheet(sheet):
  if(sheet not in wb.sheetnames):
    wb.create_sheet(sheet)
  currentsheet = wb.get_sheet_by_name(sheet)
  currentsheet.append(('Finished Good List','','',''))
  currentsheet.append(('#', 'Item Description', 'Quantity', 'Unit'))
  currentsheet.append(('1', sheet, '1', 'Pc'))
  currentsheet.append(('End of FG', '', '', ''))
  currentsheet.append(('Raw Material List', '', '', ''))
  currentsheet.append(('#', 'Item Description', 'Quantity', 'Unit'))
  wb.save(path)
  return currentsheet

def completeSheet(sheet):
  if(sheet == 'Source'):
    return False
  currentsheet = wb.get_sheet_by_name(sheet)
  currentsheet.append(('End of RM', '', '', ''))
  wb.save(path)
  return currentsheet

sheet=wb.active
for i in range(2,13):   
  offset = 2
  cell=sheet.cell(row=i,column=offset)
  itemCell = sheet.cell(row=i,column=2+1)
  shift = level[cell.value]
  # figure out parent sheet for multi level BOM
  parentCell=sheet.cell(row=i+shift['row'],column=offset+shift['column'])
  quantityCell=sheet.cell(row=i,column=offset+2)
  unitCell=sheet.cell(row=i,column=offset+3)
  # If parent sheet does not exist, create new and add values
  if(parentCell.value not in wb.sheetnames):
    createSheet(parentCell.value)
  currentsheet = wb.get_sheet_by_name(parentCell.value)
  currentsheet.append((1,itemCell.value,quantityCell.value,unitCell.value))
  wb.save(path)

# Closing all sheets created except the source sheet
for sheet in wb.sheetnames:
  completeSheet(sheet)