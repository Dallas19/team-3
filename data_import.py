#Please download these libs by: Python -m pip install <package>
import pandas
import xlrd

file_path = ("StudentResults.xlsx")

#Workbook and worksheet setup
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0) 

#row,cell
#row 0 =  column names
sheet.cell_value(0,0)


comp_pref = {}
stud_pref = {}

for r in range(sheet.nrows):
    if r == 0:
        continue
    choices = []
    
    choices.append(sheet.cell_value(r,1))
    choices.append(sheet.cell_value(r,2))
    choices.append(sheet.cell_value(r,3))
    choices.append(sheet.cell_value(r,4))
    choices.append(sheet.cell_value(r,5))
                   
    stud_pref[sheet.cell_value(r,0)] = choices

print(stud_pref)

