from src.timetable import *

students = StudentList.load('../data/timetable.xlsx', '../data/subjects.xlsx')
rhs: Student = students['류현승']

print(rhs.subjects)

# for i, (_, period) in enumerate(rhs.timetable[Week.MON]):
#     print(i, period)
