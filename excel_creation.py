import os
from openpyxl import Workbook

from datetime import datetime

wb = Workbook()
ws1 = wb.active
ws1 = wb.create_sheet("sheet1", 0)  # insert at first position
#ws1.title = 'sheet1'

ws1['A1'] = "col1"
ws1['B1'] = "col2"
ws1['C1'] = "col3"
ws1['D1'] = "col4"

print("Enter values to col1 separated by comas:")
col1_list=input().split(',')

print("Enter values to col2 separated by comas:")
col2_list=input().split(',')

divNo=len(col1_list)/len(col2_list)
print(divNo)
final=round(divNo)
print(final*col2_list)

new_list=col2_list*final
print(new_list)
for row, i in enumerate(col1):
    for row1, j in enumerate(new_list):
        column_cell = 'B'
        ws1[column_cell + str(row + 2)] = str(i)
        print(row1, row)
        if row1 > row:
            break
        column_cell = 'A'
        ws1[column_cell+str(row1+2)] = str(j)


#os.chdir(r'')
datestring = datetime.strftime(datetime.now(), '%Y_%m_%d')
wb.save('file1' + datestring + '.xlsx')
