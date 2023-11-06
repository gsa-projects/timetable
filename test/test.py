from src.preset import *

students = StudentList.from_xlsx('../data/timetable.xlsx')
rhs = students[22027]

for e in rhs.timetable.items():
    print(e)
