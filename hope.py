#!/usr/bin/python

## programe-Shapley algorithm
# http://www.nrmp.org/match-process/match-algorithm/
# http://www.nber.org/papers/w6963

# From http://rosettacode.org/wiki/Stable_marriage_problem
# Uses deepcopy to make actual copy of the contents of the dictionary in a new object
# http://pymotw.com/2/copy/

import copy
from collections import defaultdict

#Create python dictionary with names as keys, values as list
studentprefers = {'Gracey Gould': ['AT&T - Sales Rep Intern ', 'Hilton Suites - Event Planner Intern', 'Cisco - Systems Engineer Intern', 'Cisco - Software Engineer Intern', 'Wells Fargo - Accountant Intern'], 'Gloria Durham': ['Wells Fargo - Accountant Intern', 'AT&T - Sales Rep Intern ', 'Hilton Suites - Food Coordinator ', 'AT&T - Business Line Intern', 'Cisco - Cyber Security Intern'], 'Tobey Logan': ['Hilton Suites - Event Planner', 'Cisco - Software Engineering Intern', 'Cisco - Cyber Security Intern', 'AT&T - Sales Rep Intern ', 'Hilton Suites - Food Coordinator '], 'Marcel Taylor': ['AT&T - Business Intern', 'Hilton Suites - Food Coordinator ', 'AT&T - Sales Rep Intern ', 'Cisco - Systems Engineer Intern', 'Cisco - Software Engineer Intern'], 'Annette Mcclure': ['Hilton Suites - Event Planner', 'AT&T - Sales Rep Intern ', 'Cisco - Systems Engineer Intern', 'Cisco - Software Engineer Intern', 'AT&T - Business Line Intern'], 'Tamanna Howells': ['Cisco - Software Engineering Intern', 'AT&T - Business Intern', 'AT&T - Cyber Security Intern', 'AT&T - Sales Rep Intern ', 'Hilton Suites - Event Planner Intern'], 'Terri Whiteley': ['Cisco - Cyber Security Intern', 'Cisco - Software Engineer Intern', 'Hilton Suites - Food Coordinator ', 'AT&T - Sales Rep Intern ', 'Wells Fargo - Accountant Intern'], 'Zayden Firth': ['Hilton Suites - Food Coordinator ', 'AT&T - Sales Rep Intern', 'Wells Fargo - Accountant Intern', 'AT&T - Business Line Intern', 'Cisco - Software Engineer Intern'], 'Ann Escobar': ['Wells Fargo - Accountant Intern', 'Cisco - Cyber Security Intern', 'Hilton Suites - Food Coordinator ', 'Cisco - Software Engineer Intern', 'AT&T - Sales Rep Intern '], "Haleema O'Reilly": ['Cisco - Cyber Security Intern', 'AT&T - Business Intern', 'AT&T - Sales Rep Intern ', 'Hilton Suites - Food Coordinator ', 'Cisco - Software Engineer Intern']}
programprefers = {'AT&T - Sales Rep Intern ': ['Gracey Gould', 'Gloria Durham', 'Tobey Logan', 'Marcel Taylor', 'Annette Mcclure', 'Tamanna Howells', 'Terri Whiteley', 'Zayden Firth', 'Ann Escobar'], 'AT&T - Software Intern ': ['Tobey Logan', 'Terri Whiteley', "Haleema O'Reilly", 'Ann Escobar', 'Gloria Durham', 'Zayden Firth'], 'AT&T - Business Intern': ['Tamanna Howells', 'Zayden Firth', 'Terri Whiteley', 'Zayden Firth', "Haleema O'Reilly", 'Gloria Durham'], 'Cisco - Cyber Security Intern': ["Haleema O'Reilly", 'Ann Escobar', 'Marcel Taylor', 'Gloria Durham', 'Ann Escobar', 'Tobey Logan'], 'Cisco - Software Engineering Intern': ['Annette Mcclure', 'Ann Escobar', "Haleema O'Reilly", 'Zayden Firth', 'Tobey Logan', 'Gracey Gould'], 'Cisco - Systems Intern': ['Gloria Durham', 'Marcel Taylor', 'Terri Whiteley', 'Ann Escobar', 'Tamanna Howells', 'Gloria Durham'], 'Wells Fargo - Accountant Intern': ['Marcel Taylor', 'Terri Whiteley', 'Annette Mcclure', 'Ann Escobar', 'Gracey Gould', 'Tamanna Howells'], 'Hilton Suites - Event Planner': ['Zayden Firth', 'Terri Whiteley', 'Marcel Taylor', 'Tobey Logan', 'Gracey Gould', 'Ann Escobar'], 'Hilton Suites - Food Coordinator ': ['Tobey Logan', "Haleema O'Reilly", 'Annette Mcclure', 'Gracey Gould', 'Tobey Logan', 'Gloria Durham', 'Annette Mcclure']}
programSlots = {
 'AT&T - Sales Rep Intern ': 1,
 'AT&T - Software Intern ': 1,
 'AT&T - Business Intern': 1,
 'Cisco - Cyber Security Intern': 1,
 'Cisco - Software Engineering Intern': 1,
 'Cisco - Systems Intern': 2,
 'Wells Fargo - Accountant Intern': 1,
 'Hilton Suites - Event Planner': 2,
 'Hilton Suites - Food Coordinator ': 2
 }
 
students = sorted(studentprefers.keys())
programs = sorted(programprefers.keys())


 
def matchmaker():
    studentsfree = students[:]
    studentslost = []
    matched = {}
    studentsprogram =""
    for programName in programs:
        if programName not in matched:
             matched[programName] = list()
    studentprefers2 = copy.deepcopy(studentprefers)
    programprefers2 = copy.deepcopy(programprefers)
    while studentsfree:
        student = studentsfree.pop(0)
        print("%s is on the market" % (student))
        studentslist = studentprefers2[student]
        program = studentslist.pop(0)
        if student in programprefers[program]:
            print("  %s (program's #%s) is checking out %s (student's #%s)" % (student, (programprefers[program].index(student)+1), program, (studentprefers[student].index(program)+1)) )
            tempmatch = matched.get(program)
            if len(tempmatch) < programSlots.get(program):
                # Program's free
                if student not in matched[program]:
                    matched[program].append(student)
                    print("    There's a spot! Now matched: %s and %s" % (student.upper(), program.upper()))
            else:
                # The student proposes to an full program!
                programslist = programprefers2[program]
                for (i, matchedAlready) in enumerate(tempmatch):
                    if programslist.index(matchedAlready) > programslist.index(student):
                        # Program prefers new student
                        if student not in matched[program]:
                            matched[program][i] = student
                            print("  %s dumped %s (program's #%s) for %s (program's #%s)" % (program.upper(), matchedAlready, (programprefers[program].index(matchedAlready)+1), student.upper(), (programprefers[program].index(student)+1)))
                            if studentprefers2[matchedAlready]:
                                # Ex has more programs to try
                                studentsfree.append(matchedAlready)
                            else:
                                studentslost.append(matchedAlready)
                    else:
                        # Program still prefers old match
                        print("  %s would rather stay with %s (their #%s) over %s (their #%s)" % (program, matchedAlready, (programprefers[program].index(matchedAlready)+1), student, (programprefers[program].index(student)+1)))
                        if studentslist:
                            # Look again
                            studentsfree.append(student)
                        else:
                            studentslost.append(student)
    print 
    for lostsoul in studentslost:
        print('%s did not match' % lostsoul)
    return (matched, studentslost)





def check(matched):
    studentsprogram =""
    inversematched = defaultdict(list)
    for programName in matched.keys():
        for studentName in matched[programName]:
            inversematched[programName].append(studentName)

    for programName in matched.keys():
        for studentName in matched[programName]:

            programNamelikes = programprefers[programName]
            programNamelikesbetter = programNamelikes[:programNamelikes.index(studentName)]
            helikes = studentprefers[studentName]
            helikesbetter = helikes[:helikes.index(programName)]
            for student in programNamelikesbetter:
                for p in inversematched.keys():
                    if student in inversematched[p]:
                        studentsprogram = p
                studentlikes = studentprefers[student]
				
				## Not sure if this is correct
                try:
                    studentlikes.index(studentsprogram)
                except ValueError:
                    continue
				
                if studentlikes.index(studentsprogram) > studentlikes.index(programName):
                    print("%s and %s like each other better than "
                          "their present match: %s and %s, respectively"
                          % (programName, student, studentName, studentsprogram))
                    return False
            for program in helikesbetter:
                programsstudents = matched[program]
                programlikes = programprefers[program]
                for programsstudent in programsstudents:
                    if programlikes.index(programsstudent) > programlikes.index(studentName):
                        print("%s and %s like each other better than "
                              "their present match: %s and %s, respectively"
                              % (studentName, program, programName, programsstudent))
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
matched[programs[0]], matched[programs[1]] = matched[programs[1]], matched[programs[0]]
for program in programs[:2]:
    print('  %s is now matched to %s' % (program, matched[program]))
print
print('Match stability check PASSED'
      if check(matched) else 'Match stability check FAILED')
