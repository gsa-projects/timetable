from src.preset import *

students = StudentList.from_xlsx('../data/timetable.xlsx')
rhs = students[0:7]

print(rhs)
