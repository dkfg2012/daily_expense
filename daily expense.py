import openpyxl
import datetime
import time
import calendar

column = ['transport', 'breakfast', 'lunch', 'dinner', 'snack', 'drink', 'entertainment', 'others (specific)']
day_of_week = {"Sun":"B", "Mon":"C", "Tue":"D", "Wed":"E", "Thu":"F", "Fri":"G", "Sat":"H"}
wb = openpyxl.load_workbook("expense_table.xlsx")
sheet = wb['Sheet1']

current_row = 1

def today_weekday():
    today = datetime.datetime.now().strftime("%d-%m-%Y")
    weekday = tuple(list(map(int, datetime.datetime.now().strftime("%Y,%m,%d").split(','))))
    weekday = calendar.day_name[calendar.weekday(*weekday)][0:3]
    return today, weekday

today, weekday = today_weekday()

max_row = sheet.max_row
max_column = sheet.max_column

day_cell = day_of_week[weekday]

# if type(sheet["B1"].value) == type(None):
#     sheet["B1"].value = today

#seperate from last week and this week, write later
#if weekday == "Sun":


if type(sheet['A2'].value) == type(None):
    for label in range(len(column)):
        sheet['A'+str(label+2)].value = column[label]



wb.save("expense_table.xlsx")