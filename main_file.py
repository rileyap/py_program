from logging import raiseExceptions
from tracemalloc import start
from openpyxl import load_workbook
from datetime import date, datetime, time
from datetime import timedelta
from datetime import *

from scipy.fftpack import diff
path = "C:\\Users\\Lifeguard\\Desktop\\Timecards2022\\"
print("Enter your intials capitalized:")
nameInput = input()
endFile = ""
choosenKey = ""
shiftWorked = ""
startShift = ""
endShift = ""
dict = {'RP': 'PEARSON, RILEY', 'JG': 'GONZALEZ, JOSH', 'OA' :'ALBERS, OLIVIA', 'AB': 'BUBP, AUDRA', 'AF': 'FULLENKAMP, ALLYSEN', 
'CF': 'FULLENKAMP, CARSON', 'PG': 'GUGGENBILLER, PAIGE', 'KH': 'HEITKAMP, KAYLA', 'LH': 'HIPPLE, LAURA', 'JK': 'JOELLE, KAUP',
'ABK': 'KNAPKE, ABIGAIL', 'ALK': 'KNAPKE, ALLISON', 'IK': 'KNAPKE, ISAAC', 'NN': 'NGUYEN, NGOC', 'AV': 'VAUGHN, ALLISON', 
'FW': 'WENDEL, FAITH' }
dict2 = {'PEARSON, RILEY': 'prompt', 'GONZALEZ, JOSH': 'prompt', 'VAUGHN, ALLISON': 'prompt'}
dict3 = {'G': 'N', 'g': 'N', 'L': 'O', 'l': 'O', 'M': 'P', 'm': 'P'}
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
if shiftWorked == "":
    print("Enter your type of shift: G for guard, L for Lesson")
    shiftWorked = input()

for key in dict3:
    if shiftWorked == key:
        shiftWorked = dict3[key]
        
print("Enter your shift length: D for day, N for night, C for custom")

hoursWorked = input()
if hoursWorked == 'C' or hoursWorked == 'c':
    print("Enter the start of your shift: example 12:00 PM")
    startShift = input()
    print("Enter the end of your shift: example 3:00 PM")
    endShift = input()

if hoursWorked == "d" or hoursWorked == "D":
    startShift = '12:45 PM'
    endShift = '5:00 PM'
if hoursWorked == 'n' or hoursWorked == 'N':
    startShift = "5:45 PM"
    endShift = "8:15 PM"
# print(startShift, endShift)
wb = load_workbook(filename= actualFileName)
sheet = wb.active
today = date.today()

dateFromExcel = sheet["H5"].value
dateFromExcel = dateFromExcel.date()

# print(dateFromExcel)
diff_date = today - dateFromExcel
diff_date = (diff_date.days)
if diff_date < 7:
    newIndex = diff_date + 14
if diff_date > 7:
    newIndex = diff_date + 15

newIndex = str(newIndex)
firstCheck = "B"
firstCheck += newIndex

if sheet[firstCheck].value != None:
    firstCheck = "D" + newIndex
    if sheet[firstCheck].value != None:
        firstCheck = "F" + newIndex
        if sheet[firstCheck].value != None:
            firstCheck = "H" + newIndex
            if sheet[firstCheck].value != None:
                firstCheck = "J" + newIndex
charOfDate = firstCheck[0]
secondDate = chr(ord(charOfDate) + 1)
secondDate += newIndex
shiftCol = shiftWorked + newIndex

# print(firstCheck)
# print(secondDate)
sheet[firstCheck] = startShift
sheet[secondDate] = endShift




FMT = '%I:%M %p'
tdelta = datetime.strptime(endShift, FMT) 
mdelta = datetime.strptime(startShift, FMT)
kDelta = tdelta - mdelta

kDelta = str(kDelta)
(h,m,s) = kDelta.split(':')
hoursInDecimal = float(h) + (float(m) /60.0)

if sheet[shiftCol].value == None:
    sheet[shiftCol] = hoursInDecimal
else:
    val = sheet[shiftCol].value
    hoursInDecimal += val
    sheet[shiftCol] = hoursInDecimal
wb.save(filename= actualFileName)

