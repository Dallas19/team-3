#Please download these libs by: Python -m pip install <package>
import pandas
import xlrd

file_path = ("StudentResults.xlsx")
file_path2 = ("student_to_job.xlsx")

#Workbook and worksheet setup for sheet 1
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)

#Workbook and worksheet setup for sheet 2
wb2 = xlrd.open_workbook(file_path2)
sheet2 = wb2.sheet_by_index(0)

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

for c in range(sheet.ncols):
    if sheet2.cell_value(0,c) == "Company":
        print("HELL YAH")

print(stud_pref)

