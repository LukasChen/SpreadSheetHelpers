import sys
import os
import re
from openpyxl import load_workbook
from openpyxl.cell import MergedCell


starting_row = 6
end_row = 16
start_col = 5
end_col = 40

student_tasks = {}
weekdays = ['MON', 'TUE', 'WED', 'THU', 'FRI']

def read_student_tasks(student_name, full_filename):
    # Read data
    wb = load_workbook(full_filename)
    ws = wb.active
    student_tasks[student_name] = {}
    weekday = 'MON'
    for i in range(starting_row, end_row):
        weekdayHeader = ws.cell(row=i,column=start_col - 1).value;
        if weekdayHeader != None:
            weekday = weekdayHeader 


        for j in range(start_col, end_col):
            cell = ws.cell(row=i, column=j)
            if weekday not in student_tasks[student_name]:
                student_tasks[student_name][weekday] = []
            if not isinstance(cell, MergedCell):
                student_tasks[student_name][weekday].append(cell.value)


# Get src files
full_filenames = os.listdir(os.getcwd())

for file in full_filenames:
    match = re.search("\.xlsx$", file)

    if match:
        # print(file)
        filename = file.split('.')[0]
        if file != sys.argv[1].split('/')[1] and file != sys.argv[2].split('/')[1]:
            read_student_tasks(filename, file)



print(student_tasks)

output = load_workbook(sys.argv[1], data_only=True)
output_sheet = output.active
student_search_row = 6
student_search_col = 12
student_end_row = 18
summary_interval = 26
data_col = 13
student_num = 6


for i in range(0, 5):
    base_row = i * summary_interval
    weekday = weekdays[i]
    print(weekday)
    
    name = ''
    index = 0
    for j in range(base_row + student_search_row, base_row + student_end_row):
        nameCell = output_sheet.cell(row=j, column=student_search_col).value
        if nameCell != None:
            index = 0
            if nameCell in student_tasks:
                name = nameCell
            else:
                continue
        for k in range(data_col, 48):
            cell = output_sheet.cell(row=j, column=k)
            if not isinstance(cell, MergedCell):
                cell.value = student_tasks[name][weekday][index]
                index += 1
        

output.save(sys.argv[2])
