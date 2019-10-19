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

# #concatenate company name and position name
# companyColNum = None
# positionColNum = None
# for c in range(sheet.ncols):
#     if sheet2.cell_value(0,c) == "Company":
#         companyColNum = c
#     if sheet2.cellvalue(r, c) == "Position":
#         positionColNum = c
    

 
#TODO:concatenate company name and position name
studentRank = {
 'Alex':  ['American', 'Mercy', 'County', 'Mission', 'General', 'Fairview', 'Saint Mark', 'City', 'Deaconess', 'Park'],
 'Brian':  ['County', 'Deaconess', 'American', 'Fairview', 'Mercy', 'Saint Mark', 'City', 'General', 'Mission', 'Park'],
 'Cassie':  ['Deaconess', 'Mercy', 'American', 'Fairview', 'City', 'Saint Mark', 'Mission', 'Park', 'County', 'General'],
 'Dana':  ['Mission', 'Saint Mark', 'Fairview', 'Park', 'Deaconess', 'Mercy', 'General', 'City', 'County', 'American'],
 'Edward':   ['General', 'Fairview', 'City', 'County', 'Saint Mark', 'Mercy', 'American', 'Mission', 'Deaconess', 'Park'],
 'Faith': ['City', 'American', 'Fairview', 'Park', 'Mercy', 'Mission', 'County', 'General', 'Deaconess', 'Saint Mark'],
 'George':  ['Park', 'Mercy', 'Mission', 'City', 'County', 'American', 'Fairview', 'Deaconess', 'General', 'Saint Mark'],
 'Hannah':  ['American', 'Mercy', 'Deaconess', 'Saint Mark', 'Mission', 'County', 'General', 'City', 'Park', 'Fairview'],
 'Ian':  ['Park', 'County', 'Fairview', 'Deaconess', 'City', 'American', 'Saint Mark', 'Mission', 'General', 'Mercy'],
 'Jessica':  ['American', 'Saint Mark', 'General', 'Park', 'Mercy', 'City', 'Fairview', 'County', 'Mission', 'Deaconess']}

#TODO:concatenate company name and position name
companyRank = {
 'American':  ['Brian', 'Faith', 'Jessica', 'George', 'Ian', 'Alex', 'Dana', 'Edward', 'Cassie', 'Hannah'],
 'City':  ['Brian', 'Alex', 'Cassie', 'Faith', 'George', 'Dana', 'Ian', 'Edward', 'Jessica', 'Hannah'],
 'County': ['Faith', 'Brian', 'Edward', 'George', 'Hannah', 'Cassie', 'Ian', 'Alex', 'Dana', 'Jessica'],
 'Fairview':  ['Faith', 'Jessica', 'Cassie', 'Alex', 'Ian', 'Hannah', 'George', 'Dana', 'Brian', 'Edward'],
 'Mercy':  ['Jessica', 'Hannah', 'Faith', 'Dana', 'Alex', 'George', 'Cassie', 'Edward', 'Ian', 'Brian'],
 'Saint Mark':  ['Brian', 'Alex', 'Edward', 'Ian', 'Jessica', 'Dana', 'Faith', 'George', 'Cassie', 'Hannah'],
 'Park':  ['Jessica', 'George', 'Hannah', 'Faith', 'Brian', 'Alex', 'Cassie', 'Edward', 'Dana', 'Ian'],
 'Deaconess': ['George', 'Jessica', 'Brian', 'Alex', 'Ian', 'Dana', 'Hannah', 'Edward', 'Cassie', 'Faith'],
 'Mission':  ['Ian', 'Cassie', 'Hannah', 'George', 'Faith', 'Brian', 'Alex', 'Edward', 'Jessica', 'Dana'],
 'General':  ['Edward', 'Hannah', 'George', 'Alex', 'Brian', 'Jessica', 'Cassie', 'Ian', 'Faith', 'Dana']}

companySlots = {
 'American': 1,
 'City': 1,
 'County': 1,
 'Fairview': 1,
 'Mercy': 1,
 'Saint Mark': 2,
 'Park': 1,
 'Deaconess': 2,
 'Mission': 2,
 'General': 9}
 
students = list(studentRank.keys())
companys = list(companyRank.keys())
 
def matchmaker():
    unmatchedStudents = students[:]
    studentslost = []
    matched = {}
    for companyName in companys:
        if companyName not in matched:
             matched[companyName] = list()
    studentRank2 = copy.deepcopy(studentRank)
    # companyRank2 = copy.deepcopy(companyRank)
    while unmatchedStudents:
        student = unmatchedStudents.pop(0)
        print("%s is on the market" % (student))
        studentRankList = studentRank2[student]
        print(studentRankList)
        company = studentRankList.pop(0)
        print("  %s (company's #%s) is checking out %s (student's #%s)" % (student, (companyRank[company].index(student)+1), company, (studentRank[student].index(company)+1)) )
        tempmatch = matched.get(company)
        print(tempmatch)
        if len(tempmatch) < companySlots.get(company):
            if student not in matched[company]:
                matched[company].append(student)
                print("    There's a spot! Now matched: %s and %s" % (student.upper(), company.upper()))
        else:
            # The student proposes to an full company!
            companyslist = companyRank[company]
            for (i, matchedAlready) in enumerate(tempmatch):
                if companyslist.index(matchedAlready) > companyslist.index(student):
                    # company prefers new student
                    if student not in matched[company]:
                        matched[company][i] = student
                        print("  %s dumped %s (company's #%s) for %s (company's #%s)" % (company.upper(), matchedAlready, (companyRank[company].index(matchedAlready)+1), student.upper(), (companyRank[company].index(student)+1)))
                        if studentRank[matchedAlready]:
                            # Ex has more companys to try
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
                companysstudents = matched[company]
                companylikes = companyRank[company]
                for companysstudent in companysstudents:
                    if companylikes.index(companysstudent) > companylikes.index(studentName):
                        print("%s and %s like each other better than "
                              "their present match: %s and %s, respectively"
                              % (studentName, company, companyName, companysstudent))
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
matched[companys[0]], matched[companys[1]] = matched[companys[1]], matched[companys[0]]
for company in companys[:2]:
    print('  %s is now matched to %s' % (company, matched[company]))
print
print('Match stability check PASSED'
      if check(matched) else 'Match stability check FAILED')
