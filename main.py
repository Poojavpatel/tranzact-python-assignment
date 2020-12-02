import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

path = 'data.xlsx'
wb=load_workbook(path)

sheets=['Fan', 'Toy','Motor','Wires','Iron chips','Copper granule', 'Metal Tools']

level = {
  ".1" : {'row':0, 'column':-1},
  "..2" : {'row':-1, 'column':1},
  "..3" : {'row':-1, 'column':1}
}

for sheet in sheets:
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

sheet=wb.active
for i in range(2,13):   
  for j in range(2,6):
    cell=sheet.cell(row=i,column=j)
    if(j==2):
      itemCell = sheet.cell(row=i,column=j+1)
      shift = level[cell.value]
      parentCell=sheet.cell(row=i+shift['row'],column=j+shift['column'])
      quantityCell=sheet.cell(row=i,column=j+2)
      unitCell=sheet.cell(row=i,column=j+3)
      currentsheet = wb.get_sheet_by_name(parentCell.value)
      currentsheet.append((1,itemCell.value,quantityCell.value,unitCell.value))
      wb.save(path)
    
for sheet in sheets:
  currentsheet = wb.get_sheet_by_name(sheet)
  currentsheet.append(('End of RM', '', '', ''))
wb.save(path)
      
  
  