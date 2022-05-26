from openpyxl import load_workbook
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

sheet["A1"] = "l"

wb.save(filename= actualFileName)

