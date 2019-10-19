#!/usr/bin/python
 
## companye-Shapley algorithm
# http://www.nrmp.org/match-process/match-algorithm/
# http://www.nber.org/papers/w6963
 
# From http://rosettacode.org/wiki/Stable_marriage_problem
# Uses deepcopy to make actual copy of the contents of the dictionary in a new object
# http://pymotw.com/2/copy/
 
import copy
from collections import defaultdict
import pandas
import xlrd

#change to whatever files used by EIF
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

#print(studentRank)
#print(companyRank)

# #concatenate company name and position name
# companyColNum = None
# positionColNum = None
# for c in range(sheet.ncols):
#     if sheet2.cell_value(0,c) == "Company":
#         companyColNum = c
#     if sheet2.cellvalue(r, c) == "Position":
#         positionColNum = c
 
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
        print("#######Unmatched Students##########")
        print(unmatchedStudents)
        student = unmatchedStudents.pop(0)
        print("%s is on the market" % (student)) 
        studentRankList = studentRank2[student]
        #print(studentRankList)
        while studentRankList:
            company = studentRankList.pop(0)
            print(company)
            print(companyRank[company])
            if student in companyRank[company]:
                print("  %s (company's #%s) is checking out %s (student's #%s)" % (student, (companyRank[company].index(student)+1), company, (studentRank[student].index(company)+1)) )
                tempmatch = matched.get(company)
                print("tempmatch")
                print(tempmatch)
                if len(tempmatch) < companySlots.get(company):
                    if student not in matched[company]:
                        matched[company].append(student)
                        print("    There's a spot! Now matched: %s and %s" % (student.upper(), company.upper()))
                        break
                else:
                    # The student proposes to an full company!
                    companieslist = companyRank2[company]
                    print(tempmatch)
                    # [(0, 'Grace'), (1, 'Bob')] enumerate returns the index and element as a tuple from the iterable given
                    for (i, matchedAlready) in enumerate(tempmatch):
                        #Check whether the currect position match is lower ranked than the current student
                        if companieslist.index(matchedAlready) > companieslist.index(student):
                            # company prefers new student
                            if student not in matched[company]:
                                matched[company][i] = student
                                print("  %s dumped %s (company's #%s) for %s (company's #%s)" % (company.upper(), matchedAlready, (companyRank[company].index(matchedAlready)+1), student.upper(), (companyRank[company].index(student)+1)))
                                if studentRank[matchedAlready]:
                                    # Ex has more companies to try
                                    unmatchedStudents.append(matchedAlready)
                                else:
                                    studentslost.append(matchedAlready)
                        else:
                            # company still prefers old match
                            print("  %s would rather stay with %s (their #%s) over %s (their #%s)" % (company, matchedAlready, (companyRank[company].index(matchedAlready)+1), student, (companyRank[company].index(student)+1)))
                            if studentRankList:
                                # Look again
                                unmatchedStudents.append(student)
                            else:
                                studentslost.append(student)
    print
    for lostsoul in studentslost:
        print('%s did not match' % lostsoul)
    return (matched, studentslost)
 
 
 
 
 
def check(matched):
    inversematched = defaultdict(list)
    for companyName in matched.keys():
        for studentName in matched[companyName]:
            inversematched[companyName].append(studentName)
 
    for companyName in matched.keys():
        for studentName in matched[companyName]:
 
            companyNamelikes = companyRank[companyName]
            companyNamelikesbetter = companyNamelikes[:companyNamelikes.index(studentName)]
            helikes = studentRank[studentName]
            helikesbetter = helikes[:helikes.index(companyName)]
            studentscompany = ''
            for student in companyNamelikesbetter:
                for p in inversematched.keys():
                    if student in inversematched[p]:
                        studentscompany = p
                studentlikes = studentRank[student]
                               
                                ## Not sure if this is correct
                try:
                    studentlikes.index(studentscompany)
                except ValueError:
                    continue
                               
                if studentlikes.index(studentscompany) > studentlikes.index(companyName):
                    print("%s and %s like each other better than "
                          "their present match: %s and %s, respectively"
                          % (companyName, student, studentName, studentscompany))
                    return False
            for company in helikesbetter:
                companystudents = matched[company]
                companylikes = companyRank[company]
                for companystudent in companystudents:
                    if companylikes.index(companystudent) > companylikes.index(studentName):
                        print("%s and %s like each other better than "
                              "their present match: %s and %s, respectively"
                              % (studentName, company, companyName, companystudent))
                        return False
    return True
 
 
print('\nPlay-by-play:')
(matched, studentslost) = matchmaker()
 
print('\nCouples:')
print('  ' + ',\n  '.join('%s is matched to %s' % couple
                          for couple in sorted(matched.items())))
print
print('Match stability check PASSED'
      if check(matched) else 'Match stability check FAILED')
 
print('\n\nSwapping two matches to introduce an error')
matched[companies[0]], matched[companies[1]] = matched[companies[1]], matched[companies[0]]
for company in companies[:2]:
    print('  %s is now matched to %s' % (company, matched[company]))
print
print('Match stability check PASSED'
      if check(matched) else 'Match stability check FAILED')
