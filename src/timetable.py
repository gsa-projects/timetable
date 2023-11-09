import pandas as pd
import re
from datetime import date, time, timedelta, datetime
from math import *
from dataclasses import dataclass
from enum import Enum

max_period = 9  # 최대 교시 수

class Week(Enum):
	MON = 0
	TUE = 1
	WED = 2
	THU = 3
	FRI = 4
	
	def __call__(self, value: int):
		if value == 0:
			return Week.MON
		elif value == 1:
			return Week.TUE
		elif value == 2:
			return Week.WED
		elif value == 3:
			return Week.THU
		elif value == 4:
			return Week.FRI
		else:
			raise ValueError(f'invalid week: {value}')
	
	@staticmethod
	def from_string(text: str):
		if isinstance(text, Week):
			return text
		
		mon = ['월', '월요일', 'mon', 'monday']
		tue = ['화', '화요일', 'tue', 'tuesday']
		wed = ['수', '수요일', 'wed', 'wednesday']
		thu = ['목', '목요일', 'thu', 'thursday']
		fri = ['금', '금요일', 'fri', 'friday']
		sat = ['토', '토요일', 'sat', 'saturday']
		sun = ['일', '일요일', 'sun', 'sunday']
		
		text = text.lower()
		
		if text in mon:
			return Week.MON
		elif text in tue:
			return Week.TUE
		elif text in wed:
			return Week.WED
		elif text in thu:
			return Week.THU
		elif text in fri:
			return Week.FRI
		else:
			raise ValueError(f'invalid week: {text}')

class SubjectType(Enum):
	인문 = 0
	예체능 = 1
	교양 = 2
	수학 = 3
	물리학 = 4
	화학 = 5
	생명과학 = 6
	지구과학 = 7
	정보과학 = 8

	@property
	def color(self):
		if self == SubjectType.인문:
			return ['#F4B8C4', '#F09AAA']		# 대충 빨간색
		elif self == SubjectType.예체능:
			return ['#F9E4BE', '#F6D69C']		# 대충 노란색
		elif self == SubjectType.교양:
			return ['#D8E4F3', '#BCD1EA']		# 대충 하늘색
		elif self == SubjectType.수학:
			return ['#C9E8D2', '#AEDCBB']		# 대충 초록색
		else:		# 과학
			return ['#D0CBF1', '#BBB3EB']		# 대충 보라색

@dataclass
class Subject:
	name: str
	time: int
	nth: int
	type: SubjectType = None

	def __post_init__(self):
		if any(c in self.name for c in ['물리', '역학']):
			self.type = SubjectType.물리학
		elif any(c in self.name for c in ['화학']):
			self.type = SubjectType.화학
		elif any(c in self.name for c in ['생명', '생물', '생리학', '생태']):
			self.type = SubjectType.생명과학
		elif any(c in self.name for c in ['지구', '천문']):
			self.type = SubjectType.지구과학
		elif any(c in self.name for c in ['딥러닝', '프로그래밍', '알고리즘']):
			self.type = SubjectType.정보과학
		elif any(c in self.name for c in ['적분', '선형', '수학', '기하', '미분', '정수론']):
			self.type = SubjectType.수학
		elif any(c in self.name for c in ['음악', '미술', '체육', '건강']):
			self.type = SubjectType.예체능
		elif any(c in self.name for c in ['정치', '영작', '영어', '회화', '고전', '경제', '문학', '중국', '일본', '아시아', '독서', '작문']):
			self.type = SubjectType.인문
		else:
			self.type = SubjectType.교양
	
	def __hash__(self):
		return hash(f'{self.name}_{self.nth}')
	
	def __repr__(self):
		return f'{self.name} {self.nth}반 ({self.time}시수)'
	
	__str__ = __repr__

Subject.GAP = Subject('공강', 0, 0)

class SubjectSet:
	def __init__(self, args=None):
		if args is None:
			self.__content: set[Subject] = set()
		elif isinstance(args, set):
			self.__content: set[Subject] = args
		elif isinstance(args, SubjectSet):
			self.__content: set[Subject] = args.__content
		elif isinstance(args, Subject):
			self.__content: set[Subject] = {args}
		else:
			raise TypeError(f'unsupported operand type(s) for |: {type(self)} and {type(args)}')
	
	def __len__(self):
		return len(self.__content)
	
	def add(self, *args: Subject):
		self.__content.add(*args)
	
	def __or__(self, value: any):
		if isinstance(value, Student):
			return SubjectSet(self.__content | value.subjects)
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.__content | value.__content)
		elif isinstance(value, Subject):
			return SubjectSet(self.__content | {value})
		else:
			raise TypeError(f'unsupported operand type(s) for |: {type(self)} and {type(value)}')
	
	def __and__(self, value: any):
		if isinstance(value, Student):
			return self & value.subjects
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.__content & value.__content)
		elif isinstance(value, Subject):
			return SubjectSet(self.__content & {value})
		else:
			raise TypeError(f'unsupported operand type(s) for &: {type(self)} and {type(value)}')
	
	def __xor__(self, value: any):
		if isinstance(value, Student):
			return SubjectSet(self.__content ^ value.subjects)
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.__content ^ value.__content)
		elif isinstance(value, Subject):
			return SubjectSet(self.__content ^ {value})
		else:
			raise TypeError(f'unsupported operand type(s) for ^: {type(self)} and {type(value)}')
	
	def __sub__(self, value: any):
		if isinstance(value, Student):
			return SubjectSet(self.__content - value.subjects)
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.__content - value.__content)
		elif isinstance(value, Subject):
			return SubjectSet(self.__content - {value})
		else:
			raise TypeError(f'unsupported operand type(s) for -: {type(self)} and {type(value)}')
	
	def __repr__(self) -> str:
		ret = '{\n'
		for subject in self.__content:
			ret += f'  {subject}\n'
		ret += f'}} ({len(self.__content)}과목, {sum(map(lambda x: x.time, self.__content))}시수)'
		
		return ret
	
	def __iter__(self):
		return iter(self.__content)
	
	__str__ = __repr__
	
	def to_str(self) -> set[str]:
		return set(map(str, self.__content))
	
	def to_time(self) -> list[int]:
		return list(map(lambda x: x.time, self.__content))
	
	def find(self, name: str, nth=None) -> Subject | None:
		for subject in self.__content:
			if subject.name.startswith(name) and (nth is None or subject.nth == nth):
				return subject
		return None

class TimeTable:
	"""
	me.timetable[1]
	# 1교시 과목들 전부

	me.timetable['월']
	# 월요일 과목들 전부

	me.timetable['월'][1]
	# 월요일 1교시 과목

	me.timetable[1]['월']
	# 월요일 1교시 과목
	"""
	
	def __init__(self, week_range=range(0, len(Week)), period_range=range(1, max_period + 1), data=None):
		self.__content: dict[tuple[Week, int], Subject] = {}
		self.week_range = week_range
		self.period_range = period_range
		
		if isinstance(data, TimeTable):
			for week in map(Week, week_range):
				for period in period_range:
					self.__content[week, period] = data.__content[week, period]
		else:
			for week in map(Week, week_range):
				for period in period_range:
					self.__content[week, period] = Subject.GAP
					
	def value(self):
		if len(self.week_range) == 1 and len(self.period_range) == 1:
			return self.__content[Week(self.week_range.start), self.period_range.start]
		else:
			raise ValueError(f'invalid value: {self}')
	
	def items(self):
		return {key: value for key, value in self.__content.items() if value != Subject.GAP}.items()
	
	def __iter__(self):
		return iter({key: value for key, value in self.__content.items() if value != Subject.GAP})
	
	def __getitem__(self, key):
		if isinstance(key, tuple) and len(key) == 2:
			return self[key[0]][key[1]]
		else:
			if isinstance(key, int):
				return TimeTable(self.week_range, range(key, key + 1), data=self)
			elif isinstance(key, slice):
				if isinstance(key.start, str) or isinstance(key.stop, str):
					if key.start is None:
						return TimeTable(range(Week.MON.value, Week.from_string(key.stop).value), self.period_range, data=self)
					elif key.stop is None:
						return TimeTable(range(Week.from_string(key.start).value, Week.FRI.value + 1), self.period_range, data=self)
					else:
						start = Week.from_string(key.start)
						stop = Week.from_string(key.stop)
						return TimeTable(range(start.value, stop.value), self.period_range, data=self)
				elif isinstance(key.start, int) or isinstance(key.stop, int):
					if key.start is None:
						return TimeTable(self.week_range, range(1, key.stop), data=self)
					elif key.stop is None:
						return TimeTable(self.week_range, range(key.start, max_period + 1), data=self)
					else:
						start = key.start
						stop = key.stop
						return TimeTable(self.week_range, range(start, stop), data=self)
				else:
					raise TypeError(f'unsupported operand type(s) for []: {type(self)} and {type(key)}')
			elif isinstance(key, str):
				return self[Week.from_string(key)]
			elif isinstance(key, Week):
				return TimeTable(range(key.value, key.value + 1), self.period_range, data=self)
			else:
				raise TypeError(f'unsupported operand type(s) for []: {type(self)} and {type(key)}')
	
	def __setitem__(self, key, value: Subject):
		if isinstance(key, int):
			for week in map(Week, self.week_range):
				self.__content[week, key] = value
		elif isinstance(key, str):
			self.__setitem__(Week.from_string(key), value)
		elif isinstance(key, Week):
			for period in self.period_range:
				self.__content[key, period] = value
		elif isinstance(key, tuple) and len(key) == 2:
			k1, k2 = key
			
			if isinstance(k2, str) or isinstance(k2, Week):
				k1, k2 = k2, k1
			
			self.__content[Week.from_string(k1), k2] = value
		else:
			raise TypeError(f'unsupported operand type(s) for []: {type(self)} and {type(key)}')
	
	def __repr__(self) -> str:
		indent = " " * 2
		ret = indent
		
		week_name = ['월요일', '화요일', '수요일', '목요일', '금요일']
		for week in self.week_range:
			ret += week_name[week] + ' '
		ret += '\n'
		
		for period in self.period_range:
			ret += f'{period:<2}'
			for week in map(Week, self.week_range):
				subject = self.__content[week, period]
				if subject == Subject.GAP:
					ret += '  ' * 3
				else:
					ret += f'{"  " * (3 - len(subject.name))}{subject.name.replace(" ", "")[:3]} '
			ret += '\n'
	
		return ret
	
	def __eq__(self, other):
		if isinstance(other, TimeTable):
			return self.__content == other.__content and \
				self.week_range == other.week_range and \
				self.period_range == other.period_range
		else:
			return False
	
	__str__ = __repr__
	
	def to_google_cal(self, name):
		columns = ['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'All Day Event', 'Description']#, 'Location', 'Private']
		times = [
			time(8, 50),
			time(9, 50),
			time(10, 50),
			time(11, 50),
			time(13, 30),
			time(14, 30),
			time(15, 30),
			time(16, 40),
			time(17, 40)
		]

		data = []
		for (week, period), subject in self.items():
			start_date = date(2023, 8, 14 + week.value)
			end_date = date(2023, 12, 31)
			
			for d in map(lambda x: start_date + timedelta(days=x), range(0, (end_date - start_date).days, 7)):
				data.append([
					subject.name,
					d.strftime("%m/%d/%Y"),
					times[period - 1].strftime("%I:%M %p"),
					d.strftime("%m/%d/%Y"),
					(datetime.combine(datetime.min, times[period - 1]) + timedelta(minutes=50)).time().strftime("%I:%M %p"),
					False,
					f"{subject.nth}분반"
				])
		
		df = pd.DataFrame(data, columns=columns)
		df.to_csv(f'output/calendar/{name}.csv')

@dataclass
class Student:
	id: int
	name: str
	timetable: TimeTable
	subjects: SubjectSet
	
	def __hash__(self):
		return hash(self.id)
	
	def __eq__(self, other):
		if isinstance(other, int):
			return self.id == other
		elif isinstance(other, str):
			return self.name == other
		elif isinstance(other, tuple) and len(other) == 2:
			return self.id == other[0] and self.name == other[1]
		elif isinstance(other, Student):
			return self.id == other.id and \
				self.name == other.name and \
				self.subjects == other.subjects and \
				self.timetable == other.timetable
		else:
			return False
	
	def __or__(self, value: any) -> SubjectSet:
		if isinstance(value, Student):
			return SubjectSet(self.subjects | value.subjects)
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.subjects | value)
		elif isinstance(value, Subject):
			return SubjectSet(self.subjects | {value})
		else:
			raise TypeError(f'unsupported operand type(s) for |: {type(self)} and {type(value)}')
	
	def __and__(self, value: any) -> SubjectSet:
		if isinstance(value, Student):
			return SubjectSet(self.subjects & value.subjects)
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.subjects & value)
		elif isinstance(value, Subject):
			return SubjectSet(self.subjects & {value})
		else:
			raise TypeError(f'unsupported operand type(s) for &: {type(self)} and {type(value)}')
	
	def __xor__(self, value: any) -> SubjectSet:
		if isinstance(value, Student):
			return SubjectSet(self.subjects ^ value.subjects)
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.subjects ^ value)
		elif isinstance(value, Subject):
			return SubjectSet(self.subjects ^ {value})
		else:
			raise TypeError(f'unsupported operand type(s) for ^: {type(self)} and {type(value)}')
	
	def __sub__(self, value: any) -> SubjectSet:
		if isinstance(value, Student):
			return SubjectSet(self.subjects - value.subjects)
		elif isinstance(value, SubjectSet):
			return SubjectSet(self.subjects - value)
		elif isinstance(value, Subject):
			return SubjectSet(self.subjects - {value})
		else:
			raise TypeError(f'unsupported operand type(s) for -: {type(self)} and {type(value)}')
	
	def __repr__(self):
		return f'{self.id} {self.name}'
	
	__str__ = __repr__
	
	def get_google_cal(self):
		self.timetable.to_google_cal(self.name)

class StudentList:
	@staticmethod
	def from_xlsx(path: str):
		margin = 1  # 학생 간 줄 간격
		line_per_student = 1 + max_period + margin  # 각 학생이 차지하는 행 수
		
		# xlsx = next(iter(files.upload()))
		data = pd.read_excel(path, header=None)  # 원본 데이터
		lack_count = line_per_student - len(data) % line_per_student  # 부족한 행 수
		for _ in range(lack_count):
			data.loc[len(data)] = [nan] * 5  # 부족한 행 수만큼 빈 행 추가
		
		headcount = len(data) // line_per_student  # 학생 수
		
		students = StudentList()
		
		for i in range(headcount):
			# 각 학생 시간표 행 시작점과 끝점
			startLine = (line_per_student * i)
			endLine = line_per_student * (i + 1) - margin
			
			# 각 학생의 정보
			header = data.iloc[startLine][0]
			each_student = (int(header[:5]), header[5:])
			
			# 각 학생의 시간표
			timetable = TimeTable()
			xlsx_data = data[startLine + 1:endLine] \
				.rename(columns={0: '월', 1: '화', 2: '수', 3: '목', 4: '금'}) \
				.fillna('공강')
			xlsx_data.index = range(1, len(xlsx_data) + 1)
			
			# 각 학생의 수업
			subjects = SubjectSet()
			for period in range(max_period):
				for week, subject in xlsx_data.iloc[period].items():
					if subject != '공강':
						matched = re.match('(?P<name>[\w\d\(\) ]+) (?P<nth>\d)반 \((?P<time>\d)시간\)', subject)
						subject_instance = Subject(
							name=matched.group('name'),
							time=int(matched.group('time')),
							nth=int(matched.group('nth'))
						)
						
						subjects.add(subject_instance)
						timetable[week, period + 1] = subject_instance
			
			# 저장
			students[each_student[0]] = Student(
				id=each_student[0],
				name=each_student[1],
				timetable=timetable,
				subjects=subjects
			)
		
		return students
	
	def __init__(self, args=None):
		if args is None:
			args = []
		
		self.__students: list[Student] = list(args)
		if self.__students:
			self.__students.sort(key=lambda x: x.id)
	
	def __iter__(self):
		return iter(self.__students)
	
	def __len__(self):
		return len(self.__students)
	
	def __getitem__(self, key: slice | int | str | tuple[int, str] | Student) -> Student | None:
		if isinstance(key, slice):
			# fixme: 버그 개많음
			#  - 중간에 휴학/자퇴한 학번 있으면 그거 무시해줘야함
			indices = range(0 if key.start is None else key.start, key.stop, 1 if key.step is None else key.step)
			return StudentList([self[i] for i in indices])
		elif isinstance(key, int) and -len(self.__students) <= key < len(self.__students):
			return self.__students[key]
		else:
			if isinstance(key, str) and key.isdigit():
				return self[int(key)]

			for e in self.__students:
				if e == key:
					return e
			return None
	
	def append(self, value: Student):
		idx = len(self.__students)
		for i in range(len(self.__students)):
			if self.__students[i].id > value.id:
				idx = i
		
		self.__students.insert(idx, value)
	
	def __setitem__(self, key: int | str | tuple[int, str] | Student, value: Student):
		e = self[key]
		
		if e not in self.__students:
			self.append(value)
		else:
			self.__students[self.__students.index(e)] = value
	
	def __repr__(self) -> str:
		ret = '[\n'
		for student in self.__students:
			ret += f'  {student}\n'
		ret += f'] ({len(self.__students)}명)'
		
		return ret
	
	__str__ = __repr__
	
	def to_str(self) -> list[str, ...]:
		return list(map(str, self.__students))

@dataclass
class Ranking:
	target: Student
	subjects: SubjectSet
	score: int
