from openpyxl import load_workbook
from datetime import date, datetime, time
from datetime import timedelta
path = "C:\\git_workspaces\\py_program\\"
names =  ["RP", "JG","AK"]
print("Enter your intials capitalized:")
nameInput = input()
endFile = ""
for name in names:
    if name == nameInput:
        endFile = nameInput + "book.xlsx"
actualFileName = path + endFile

wb = load_workbook(filename= actualFileName)
sheet = wb.active
today = date.today()
sheet["A1"] = "l"
dateFromExcel = sheet["E1"].value
dateFromExcel = dateFromExcel.date()
print(dateFromExcel)
diff_date = today - dateFromExcel
print(diff_date)
wb.save(filename= actualFileName)

