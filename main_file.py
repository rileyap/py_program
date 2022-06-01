from logging import raiseExceptions
from openpyxl import load_workbook
from datetime import date, datetime, time
from datetime import timedelta
path = "C:\\Users\\Lifeguard\\Desktop\\Timecards2022\\"
print("Enter your intials capitalized:")
nameInput = input()
endFile = ""
choosenKey = ""
shiftWorked = ""
dict = {'RP': 'PEARSON, RILEY', 'JG': 'GONZALEZ, JOSH', 'OA' :'ALBERS, OLIVIA', 'AB': 'BUBP, AUDRA', 'AF': 'FULLENKAMP, ALLYSEN', 
'CF': 'FULLENKAMP, CARSON', 'PG': 'GUGGENBILLER, PAIGE', 'KH': 'HEITKAMP, KAYLA', 'LH': 'HIPPLE, LAURA', 'JK': 'JOELLE, KAUP',
'ABK': 'KNAPKE, ABIGAIL', 'ALK': 'KNAPKE, ALLISON', 'IK': 'KNAPKE, ISAAC', 'NN': 'NGUYEN, NGOC', 'AV': 'VAUGHN, ALLISON', 
'FW': 'WENDEL, FAITH' }
dict2 = {'PEARSON, RILEY': 'prompt', 'GONZALEZ, JOSH': 'prompt', 'VAUGHN, ALLISON': 'prompt'}
for key in dict:
    if key == nameInput:
        endFile = dict[key] + ".xlsx"
        choosenKey = dict[key]
actualFileName = path + endFile
if choosenKey == "":
    print("Invalid user entered. Exiting program now")
    exit()
else:
    for key in dict2:
        if key == choosenKey:
            print("Enter your type of shift: M for managment, G for guard, L for lesson")
            shiftWorked = input()


wb = load_workbook(filename= actualFileName)
sheet = wb.active
today = date.today()
sheet["A1"] = "l"
dateFromExcel = sheet["H5"].value
dateFromExcel = dateFromExcel.date()
print(dateFromExcel)
diff_date = today - dateFromExcel
print(diff_date)
wb.save(filename= actualFileName)

