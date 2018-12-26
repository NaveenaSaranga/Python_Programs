import os
from openpyxl import Workbook

from datetime import datetime

wb = Workbook()

ws2 = wb.active
ws2 = wb.create_sheet("mysheet", 0)  # insert at first position
ws2.title = 'validation_steps'
for dim in ws2.column_dimensions.values():
    dim.bestFit = True
ws2['A1'] = "#"
ws2['B1'] = "steps"
ws2['C1'] = "role"

ws2['A2'] = 1
ws2['B2'] = "Installation > Installation > verify revision is latest"
ws2['C2'] = "system-support"

ws2['A3'] = 2
ws2['B3'] = "If not latest, ask for upgrade"
ws2['C3'] = "System-admin"

ws2['A4'] = 3
ws2['B4'] = "Reset Sequence"
ws2['C4'] = "system-support"

ws2['A5'] = 4
ws2['B5'] = "Manage jobs > Drop all jobs"
ws2['C5'] = "system-support"

ws2['A6'] = 5
ws2['B6'] = "Manage jobs > Drop all jobs"
ws2['C6'] = "system-support"

ws2['A7'] = 6
ws2['B7'] = "Sync > Stop Sync"
ws2['C7'] = "system-support"

ws2['A8'] = 7
ws2['B8'] = "Sync > Reset Target"
ws2['C8'] = "system-support"

ws2['A9'] = 8
ws2['B9'] = "Sync > Start Sync"
ws2['C9'] = "system-support"
ws2['C9'] = "system-support"

ws2['A10'] = 9
ws2['B10'] = "Install >Connections> Validate DB Refresh"
ws2['C10'] = "system-support"


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
