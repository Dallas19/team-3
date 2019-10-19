#!/usr/bin/python
 
## companye-Shapley algorithm
# http://www.nrmp.org/match-process/match-algorithm/
# http://www.nber.org/papers/w6963
 
import copy
from collections import defaultdict
import pandas
import xlrd

#change to files used by EIF
studentrank_file = ("StudentResults.xlsx")
companyrank_file = ("student_to_job.xlsx")

wb_student = xlrd.open_workbook(studentrank_file)
sheet_student = wb_student.sheet_by_index(0)

wb_company = xlrd.open_workbook(companyrank_file)
sheet_company = wb_company.sheet_by_index(0)

companyRank = {}
studentRank = {}
companySlots = {}

for r in range(sheet_student.nrows):
    studentChoices = []
    if r == 0:
        continue
    student = str(sheet_student.cell_value(r,0))
    student.strip()
    for c in range(1,6):
        comp = sheet_student.cell_value(r,c)
        studentChoices.append(comp)

    studentRank[student] = studentChoices


for r in range(sheet_company.nrows):
    companyChoices = []
    company = sheet_company.cell_value(r,0)
    if r == 0:
        continue
    for c in range(sheet_company.ncols):
        student = sheet_company.cell_value(r,c)
        student.strip()
        if c == 0:
            continue
        if student == "":
            break

        companyChoices.append(student)

    companySlots[company] = 1
    companyRank[company] = companyChoices
 
students = list(studentRank.keys())
companies = list(companyRank.keys())
 
def matchmaker():
    unmatchedStudents = students[:]
    #print(unmatchedStudents)
    studentslost = []
    matched = {}
    for companyName in companies:
        if companyName not in matched:
             matched[companyName] = list()
    studentRank2 = copy.deepcopy(studentRank)
    companyRank2 = copy.deepcopy(companyRank)
    while unmatchedStudents:
        student = unmatchedStudents.pop(0)
        print("%s is available" % (student)) 
        studentRankList = studentRank2[student]
        while studentRankList:
            company = studentRankList.pop(0)
            if student in companyRank[company]:
                print("  %s (company's #%s) is checking out %s (student's #%s)" % (student, (companyRank[company].index(student)+1), company, (studentRank[student].index(company)+1)) )
                #get list of tentative matches for current company
                tempmatch = matched.get(company)
                #check if there's still a slot open and add student to matched list
                if len(tempmatch) < companySlots.get(company):
                    if student not in matched[company]:
                        matched[company].append(student)
                        print("    There's a spot! Now matched: %s and %s" % (student.upper(), company.upper()))
                        break
                #when no slot open, check whether the current student is higher ranked than the students in tempmatch
                else:
                    companieslist = companyRank2[company]
                    # [(0, 'Grace'), (1, 'Bob')] enumerate returns the index and element as a tuple from the iterable given
                    for (i, matchedAlready) in enumerate(tempmatch):
                        #Compare current student and old student rank
                        if companieslist.index(matchedAlready) > companieslist.index(student):
                            # when company prefers current student
                            if student not in matched[company]:
                                matched[company][i] = student
                                print("  %s dumped %s (company's #%s) for %s (company's #%s)" % (company.upper(), matchedAlready, (companyRank[company].index(matchedAlready)+1), student.upper(), (companyRank[company].index(student)+1)))
                                # check if old student has more companies to try 
                                if studentRank2[matchedAlready]:
                                    unmatchedStudents.append(matchedAlready)
                                else:
                                    studentslost.append(matchedAlready)
                        else:
                            # when company still prefers old match and current student has no more companies to try
                            print("  %s would rather stay with %s (their #%s) over %s (their #%s)" % (company, matchedAlready, (companyRank[company].index(matchedAlready)+1), student, (companyRank[company].index(student)+1)))
                            if not studentRankList:
                                studentslost.append(student)
    for lostsoul in studentslost:
        print('%s did not match' % lostsoul)
    return (matched, studentslost)
 
 
 #go through studentslost array and see if they are highly ranked than current students matched by a company
 #match even if the student hadn't ranked the company
def check(matched, lostSouls):
    
    modified = False

    while True:
        for company in matched.keys():
            modified = False
            for i in range(len(lostSouls)):
                if lostSouls[i] in companyRank[company]:
                    if companyRank[company].index(lostSouls[i]) < companyRank[company].index(matched[company][0]):
                        temp = matched[company][0]
                        matched[company] = [lostSouls[i]]
                        lostSouls[i] = temp
                        modified = True
        if not modified:
                break


(matched, studentslost) = matchmaker()
# Check results and modify if needed
if studentslost:
    check(matched, studentslost)

#Print all the matches
toPrintMatched = copy.deepcopy(matched)
for key in toPrintMatched.keys():
    toPrintMatched[key] = matched[key][0]
print('\nMatches:')
print('  ' + ',\n  '.join('%s is matched to %s' % couple
                          for couple in sorted(toPrintMatched.items())))

# Append the student rank and the company rank to the dictionary values
# Ex dictionary entry: { 'CompanyName - PositionName': ['StudentName', StudentRank, CompanyRank] }
for company in matched.keys():
    student = matched[company][0]
    cr = companyRank[company].index(student)+1
    if company in studentRank[student]:
        sr = studentRank[student].index(company)+1
    else:
        sr = 0
    matched[company].append(sr)
    matched[company].append(cr) 

#The matched dictionary will be used to output data in the UI
