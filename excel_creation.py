import os
from openpyxl import Workbook

from datetime import datetime

wb = Workbook()
ws1 = wb.active
ws1 = wb.create_sheet("sheet1", 0)  # insert at first position
ws1.title = 'Validation_status'
for dim in ws1.column_dimensions.values():
    dim.bestFit = True
ws1['A1'] = "Assigned POC"
ws1['B1'] = "DBs"
ws1['C1'] = "Appdb Validation"
ws1['D1'] = "Comments"

print("Enter db names separated by comas:")
db_list=input().split(',')

print("Enter Assignee names separated by comas:")
name_list=input().split(',')

divNo=len(db_list)/len(name_list)
print(divNo)
final=round(divNo)
print(final*name_list)

new_names=name_list*final
print(new_names)
for row, i in enumerate(db_list):
    for row1, j in enumerate(new_names):
        column_cell = 'B'
        ws1[column_cell + str(row + 2)] = str(i)
        print(row1, row)
        if row1 > row:
            break
        column_cell = 'A'
        ws1[column_cell+str(row1+2)] = str(j)


os.chdir(r'C:\\Users\\nsaranga\\Documents\\AppDb')
datestring = datetime.strftime(datetime.now(), '%Y_%m_%d')
wb.save('Refresh_Validation' + datestring + '.xlsx')
