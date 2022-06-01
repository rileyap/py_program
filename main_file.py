from openpyxl import load_workbook
from datetime import date, datetime, time
from datetime import timedelta
path = "C:\\git_workspaces\\py_program\\"
names =  ["RP", "JG","AK"]
print("Enter your intials capitalized:")
nameInput = input()
endFile = ""
choosenKey = ""
dict = {'RP': 'PEARSON, RILEY', 'JG': 'GONZO, JOSH'}
dict2 = {'PEARSON, RILEY': 'prompt', 'GONZO, JOSH': 'prompt', 'VAUGHN, ALLI': 'prompt'}
for key in dict:
    if key == nameInput:
        endFile = dict[key] + ".xlsx"
        choosenKey = dict[key]
actualFileName = path + endFile
for key in dict2:
    if key == choosenKey:
        print("ok")

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

