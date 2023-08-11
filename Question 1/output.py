### QUESTION-1 ###

from main import *

### OUTPUT ###

if status == pywraplp.Solver.OPTIMAL:

    student = {}
    course = {}
    days = {}
    multipleexams= {}

    print('\nStudents schedule -')
    for s in range(num_students):
        for d in range(num_days):
            for t in range(num_examslots):
                for c in range(num_courses):
                    if round(StudentExam[s,c,t,d].solution_value(), 2) == 1:
                        if S[s] in student.keys():
                            student[S[s]] += [f'{D[d]} ({T[t]}) - {C[c]}']
                        else:
                            student[S[s]] = [f'{D[d]} ({T[t]}) - {C[c]}']
    
        print(f'{S[s]} - {student[S[s]]}')
    
    print('\nCourses schedule -')
    for c in range(num_courses):
        for t in range(num_examslots):
            for d in range(num_days):
                if round(Exam[c,t,d].solution_value(), 2) == 1:
                    course[C[c]] = [f'{D[d]} ({T[t]})']
    
        print(f'{C[c]} - {course[C[c]]}')

    print('\nDays and slots schedule -')
    for d in range(num_days):
        for t in range(num_examslots):
            for c in range(num_courses):
                if round(Exam[c,t,d].solution_value(), 2) == 1:
                    if f'{D[d]} ({T[t]})' in days.keys():
                        days[f'{D[d]} ({T[t]})'] += [f'{C[c]}']
                    else:
                        days[f'{D[d]} ({T[t]})'] = [f'{C[c]}']
            if f'{D[d]} ({T[t]})' in days.keys():
                print(f'{D[d]} ({T[t]}) -',days[f'{D[d]} ({T[t]})'])

    if obj != 0:
        print('\nStudents having multiple exams on the same day -')
        for s in range(num_students):
            for d in range(num_days):
                count = 0
                for c in range(num_courses):
                    for t in range(num_examslots):
                        if round(StudentExam[s,c,t,d].solution_value(), 2) == 1:
                            count += 1
                if count == 2:
                    if S[s] in multipleexams.keys():
                        multipleexams[S[s]] += [D[d]]
                    else:
                        multipleexams[S[s]] = [D[d]]
            if S[s] in multipleexams.keys():
                print(f'{S[s]} - {multipleexams[S[s]]}')

    print('\nTotal penalty -',obj,'\n')

    # Saving data to a new worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet['A1'] = 'Exam Allocation'
    
    sheet['A3'] = 'Penalty'
    sheet['B3'] = obj

    for d in range(0,num_days):
        #sheet.cell(row = 5, column = 2+2*d, value = D[d])
        start_col = chr(ord('B') + 2*d)
        end_col = chr(ord('B') + 2*d + num_examslots - 1)
        start_cell = f'{start_col}{5}'
        end_cell = f'{end_col}{5}'
        sheet.merge_cells(f'{start_cell}:{end_cell}')
        merged_cell = sheet[start_cell]
        merged_cell.value = D[d]
        merged_cell.alignment = Alignment(horizontal='center', vertical='center')
           
    for d in range(0,num_days):
        for t in range(num_examslots):
            sheet.cell(row = 6, column = 2+t+2*d, value = T[t])
    
    for s in range(num_students):
        sheet.cell(row = 7+s, column = 1, value = S[s])

    for s in range(num_students):
        for d in range(num_days):
            for t in range(num_examslots):
                for c in range(num_courses):
                    if round(StudentExam[s,c,t,d].solution_value(), 2) == 1:
                        sheet.cell(row = 7+s, column = 2+2*d+t, value = C[c])

    workbook.save('output.xlsx')
    os.startfile('output.xlsx') # opens the output excel file

else:
    print('The problem does not have an optimal solution.')