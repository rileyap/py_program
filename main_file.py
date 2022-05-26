from openpyxl import load_workbook
path = ""
names =  ["name", "name2"]
endFile = ""
for name in names:
    continue
actualFileName = path + endFile

wb = load_workbook(filename= actualFileName)
sheet = wb.active

sheet["A1"] = "l"

wb.save(filename= actualFileName)

