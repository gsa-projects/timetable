from src.timetable import *

students = StudentList.load(
    '../data/2학년 학생별 시간표.xlsx',
    '../data/과목별 다교사수업.xlsx')

rhs = students['류현승']
parkjj = students['박진재']

print(rhs.subjects & parkjj.subjects)
