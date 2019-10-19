#Please download these libs by: Python -m pip install <package>
import xlrd
import copy
from collections import defaultdict

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

    choices.append(str(sheet.cell_value(r,1)))
    choices.append(str(sheet.cell_value(r,2)))
    choices.append(str(sheet.cell_value(r,3)))
    choices.append(str(sheet.cell_value(r,4)))
    choices.append(str(sheet.cell_value(r,5)))

    stud_pref[str(sheet.cell_value(r,0))] = choices

companySlots = {}

for r in range(sheet2.nrows):
    if r == 0:
        continue
    choices = []
    for c in range(sheet2.ncols):
        if c == 0:
            continue
        if sheet2.cell_value(r,c) == "":
            break

        choices.append(str(sheet2.cell_value(r,c)))

    comp_pref[str(sheet2.cell_value(r,0))] = choices
    companySlots[str(sheet2.cell_value(r,0))] = 1

#print(stud_pref)
print(comp_pref)

##companySlots = {
## 'American': 1,
## 'City': 1,
## 'County': 1,
## 'Fairview': 1,
## 'Mercy': 1,
## 'Saint Mark': 2,
## 'Park': 1,
## 'Deaconess': 2,
## 'Mission': 2,
## 'General': 9}


students = list(stud_pref.keys())
companys = list(comp_pref.keys())


########################

def matchmaker():
    unmatchedStudents = students[:]
    studentslost = []
    matched = {}
    tempmatch = []
    for companyName in companys:
        if companyName not in matched:
             matched[companyName] = list()
    studentRank2 = copy.deepcopy(stud_pref)
    companyRank2 = copy.deepcopy(comp_pref)
    
    while unmatchedStudents:
        student = unmatchedStudents.pop(0)
        print("%s is on the market" % (student))
        studentRankList = studentRank2[student]
        
        company = studentRankList.pop(0)
        #print(comp_pref[company])

        if student in comp_pref[company]:
            print("YE")
            print(company)
            print(comp_pref[company])
            print("  %s (company's #%s) is checking out %s (student's #%s)" % (student, (comp_pref[company].index(student)+1), company, (comp_pref[student].index(company)+1)) )
            tempmatch = matched.get(company)
            print(tempmatch)

            if len(tempmatch) < companySlots.get(company):
                if student not in matched[company]:
                    matched[company].append(student)
                    print("    There's a spot! Now matched: %s and %s" % (student.upper(), company.upper()))

        else:
            # The student proposes to an full company!
            companyslist = companyRank2[company]
            for (i, matchedAlready) in enumerate(tempmatch):
                if companyslist.index(matchedAlready) > companyslist.index(student):
                    # company prefers new student
                    if student not in matched[company]:
                        matched[company][i] = student
                        print("  %s dumped %s (company's #%s) for %s (company's #%s)" % (company.upper(), matchedAlready, (comp_pref[company].index(matchedAlready)+1), student.upper(), (comp_pref[company].index(student)+1)))
                        if comp_pref[matchedAlready]:
                            # Ex has more companys to try
                            unmatchedStudents.append(matchedAlready)
                        else:
                            studentslost.append(matchedAlready)
                else:
                    # company still prefers old match
                    print("  %s would rather stay with %s (their #%s) over %s (their #%s)" % (company, matchedAlready, (comp_pref[company].index(matchedAlready)+1), student, (comp_pref[company].index(student)+1)))
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

            companyNamelikes = comp_pref[companyName]
            companyNamelikesbetter = companyNamelikes[:companyNamelikes.index(studentName)]
            helikes = stud_pref[studentName]
            helikesbetter = helikes[:helikes.index(companyName)]
            for student in companyNamelikesbetter:
                for p in inversematched.keys():
                    if student in inversematched[p]:
                        studentscompany = p
                studentlikes = stud_pref[student]

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
                companylikes = comp_pref[company]
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
