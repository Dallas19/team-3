#!/usr/bin/python
 
## positione-Shapley algorithm
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
positionrank_file = ("student_to_job.xlsx")

wb_student = xlrd.open_workbook(studentrank_file)
sheet_student = wb_student.sheet_by_index(0)

wb_position = xlrd.open_workbook(positionrank_file)
sheet_position = wb_position.sheet_by_index(0)

positionRank = {}
studentRank = {}
positionSlots = {}

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


for r in range(sheet_position.nrows):
    positionChoices = []
    position = sheet_position.cell_value(r,0)
    if r == 0:
        continue
    for c in range(sheet_position.ncols):
        student = sheet_position.cell_value(r,c)
        student.strip()
        if c == 0:
            continue
        if student == "":
            break

        positionChoices.append(student)

    positionSlots[position] = 1
    positionRank[position] = positionChoices

#print(studentRank)
#print(positionRank)

# #concatenate position name and position name
# positionColNum = None
# positionColNum = None
# for c in range(sheet.ncols):
#     if sheet2.cell_value(0,c) == "position":
#         positionColNum = c
#     if sheet2.cellvalue(r, c) == "Position":
#         positionColNum = c
 
students = list(studentRank.keys())
positions = list(positionRank.keys())
 
def matchmaker():
    unmatchedPositions = positions[:]
    #print(unmatchedStudents)
    positionslost = []
    matched = {}
    for positionName in positions:
        if positionName not in matched:
             matched[positionName] = list()
    studentRank2 = copy.deepcopy(studentRank)
    positionRank2 = copy.deepcopy(positionRank)
    while unmatchedPositions:
        print("#######Unmatched Positions##########")
        print(unmatchedPositions)
        position = unmatchedPositions.pop(0)
        print("%s is available" % (position)) 
        positionRankList = positionRank2[position]
        #Go through all the position's ranked students
        while positionRankList:
            student = positionRankList.pop(0)
            print(student)
            print(studentRank[student])
            if position in studentRank[student]:
                print("  %s (student's #%s) is fit for %s (position's #%s)" % (position, (studentRank[student].index(position)+1), student, (positionRank[position].index(student)+1)) )
                #get list of tentative matches for current position
                tempmatch = matched.get(position)
                #check if there's still a slot open and add student to matched list
                if len(tempmatch) < positionSlots.get(position):
                    if student not in matched[position]:
                        matched[position].append(student)
                        print("    There's a spot! Now matched: %s and %s" % (student.upper(), position.upper()))
                        break
                #when no slot open, check whether the current student is higher ranked than the students in tempmatch
                else:
                    positionslist = positionRank2[position]
                    # [(0, 'Grace'), (1, 'Bob')] enumerate returns the index and element as a tuple from the iterable given
                    for (i, matchedAlready) in enumerate(tempmatch):
                        #Compare current student and old student rank
                        if positionslist.index(matchedAlready) > positionslist.index(student):
                            # when position prefers current student
                            if student not in matched[position]:
                                matched[position][i] = student
                                print("  %s dumped %s (position's #%s) for %s (position's #%s)" % (position.upper(), matchedAlready, (positionRank[position].index(matchedAlready)+1), student.upper(), (positionRank[position].index(student)+1)))
                                # check if old student has more positions to try 
                                if studentRank2[matchedAlready]:
                                    unmatchedStudents.append(matchedAlready)
                                else:
                                    studentslost.append(matchedAlready)
                        else:
                            # when position still prefers old match and current student has no more positions to try
                            print("  %s would rather stay with %s (their #%s) over %s (their #%s)" % (position, matchedAlready, (positionRank[position].index(matchedAlready)+1), student, (positionRank[position].index(student)+1)))
                            if not studentRankList:
                                studentslost.append(student)
    for lostsoul in studentslost:
        print('%s did not match' % lostsoul)
    return (matched, studentslost)
 
 
 
 
 
def check(matched):
    inversematched = defaultdict(list)
    for positionName in matched.keys():
        for studentName in matched[positionName]:
            inversematched[positionName].append(studentName)
 
    for positionName in matched.keys():
        for studentName in matched[positionName]:
 
            positionNamelikes = positionRank[positionName]
            positionNamelikesbetter = positionNamelikes[:positionNamelikes.index(studentName)]
            helikes = studentRank[studentName]
            helikesbetter = helikes[:helikes.index(positionName)]
            studentsposition = ''
            for student in positionNamelikesbetter:
                for p in inversematched.keys():
                    if student in inversematched[p]:
                        studentsposition = p
                studentlikes = studentRank[student]
                               
                                ## Not sure if this is correct
                try:
                    studentlikes.index(studentsposition)
                except ValueError:
                    continue
                               
                if studentlikes.index(studentsposition) > studentlikes.index(positionName):
                    print("%s and %s like each other better than "
                          "their present match: %s and %s, respectively"
                          % (positionName, student, studentName, studentsposition))
                    return False
            for position in helikesbetter:
                positionstudents = matched[position]
                positionlikes = positionRank[position]
                for positionstudent in positionstudents:
                    if positionlikes.index(positionstudent) > positionlikes.index(studentName):
                        print("%s and %s like each other better than "
                              "their present match: %s and %s, respectively"
                              % (studentName, position, positionName, positionstudent))
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
matched[positions[0]], matched[positions[1]] = matched[positions[1]], matched[positions[0]]
for position in positions[:2]:
    print('  %s is now matched to %s' % (position, matched[position]))
print
print('Match stability check PASSED'
      if check(matched) else 'Match stability check FAILED')
