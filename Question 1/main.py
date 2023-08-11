### QUESTION-1 ###

import math
import numpy as np
import pandas as pd
import time
from datetime import datetime

from ortools.linear_solver import pywraplp

import openpyxl
from openpyxl.styles import Alignment
import os

# MODEL FORMULATION USING OR TOOLS

### SETS AND INDICES ###

num_courses = 10
num_students = 20
num_examslots = 2
num_days = 7

C = [f'Course {c+1}' for c in range(num_courses)]
S = [f'Student {s+1}' for s in range(num_students)]         
T = ['Morning', 'Afternoon']
D = [f'Day {d+1}' for d in range(num_days)]                

### PARAMETERS ###

# Course to student allocation
def generate_binary_matrix(rows, cols, min_ones, max_ones):
    matrix = np.zeros((rows, cols), dtype=int)
    
    for col in range(cols):
        num_ones = np.random.randint(min_ones, max_ones + 1)
        ones_indices = np.random.choice(rows, num_ones, replace=False)
        matrix[ones_indices, col] = 1
        
    return matrix

min_subjects = 3
max_subjects = 5

CS = generate_binary_matrix(num_courses, num_students, min_subjects, max_subjects)

# Number of classrooms
NumClassrooms = 4

### FORMULATION ###

# START TIME
start = time.time()
now = datetime.now()
print("\nStarted at -",now.strftime("%H:%M:%S"))

solver = pywraplp.Solver.CreateSolver('SCIP')
time_limit = 30 # in minutes
solver.SetTimeLimit(1000*60*time_limit) 

### DECISION VARIABLES ### 

Exam = {(c, t, d): solver.IntVar(0, 1, f'Exam_{c}_{t}_{d}') for c in range(num_courses) for t in range(num_examslots) for d in range(num_days)}
StudentExam = {(s, c, t, d): solver.IntVar(0, 1, f'StudentExam_{s}_{c}_{t}_{d}') for s in range(num_students) for c in range(num_courses) for t in range(num_examslots) for d in range(num_days)}
pen = {(s, d): solver.IntVar(0, 1, f'pen_{s}_{d}') for s in range(num_students) for d in range(num_days)}

### CONSTRAINTS ###

# 1 - Each course can have only one exam
for c in range(num_courses):
    solver.Add(solver.Sum([Exam[c,t,d] for t in range(num_examslots) for d in range(num_days)]) == 1)

# 2 - Maximum number of exams at the same time can’t exceed the number of classrooms
for t in range(num_examslots):
    for d in range(num_days):
        solver.Add(solver.Sum([Exam[c,t,d] for c in range(num_courses)]) <= NumClassrooms)

# 3 - Each student can give maximum one exam in each slot per day
for s in range(num_students):
    for t in range(num_examslots):
        for d in range(num_days):
            solver.Add(solver.Sum([StudentExam[s,c,t,d] for c in range(num_courses)]) <= 1)

# 4 - Choosing which slot on which day a student can give the exam
for s in range(num_students):
    for c in range(num_courses):
        for t in range(num_examslots):
            for d in range(num_days):
                solver.Add(StudentExam[s,c,t,d] == CS[c][s]*Exam[c,t,d])

# 5 - Penalty if a student has to give more than one exam on the same day
for s in range(num_students):
    for d in range(num_days):
        solver.Add(solver.Sum([StudentExam[s,c,t,d] for c in range(num_courses) for t in range(num_examslots)]) <= 1 + pen[s,d])

### OBJECTIVE ###

objective_function = [pen[s,d] for s in range(num_students) for d in range(num_days)] # Minimise (penalty)

### SOLVER ###

solver.Minimize(solver.Sum(objective_function))
status = solver.Solve()
print('OPTIMAL =',status == pywraplp.Solver.OPTIMAL)

if status == pywraplp.Solver.OPTIMAL:
    """
    print('Objective value =', solver.Objective().Value())
    print('Total number of variables =', solver.NumVariables())
    print('Total number of constraints =', solver.NumConstraints())
    print('Problem solved in %f milliseconds' % solver.wall_time())
    print('Problem solved in %d iterations' % solver.iterations())
    print('Problem solved in %d branch-and-bound nodes' % solver.nodes(),'\n')"""

    obj = round(solver.Objective().Value(),2)
    print('\nTotal penalty -',obj,'\n')

else:
    print('The problem does not have an optimal solution.')

# END TIME
now = datetime.now()
print("Ended at -",now.strftime("%H:%M:%S"))
end = time.time()
print(f"Runtime of the program is {round(((end - start)/60),2)} minutes\n")