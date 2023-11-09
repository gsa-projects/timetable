from src.timetable import *

students = StudentList.from_xlsx('../data/timetable.xlsx')
rhs = students['설채환']

for subject in rhs.subjects:
    print(subject.name, subject.type)
