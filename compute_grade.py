import xlrd

info = xlrd.open_workbook('ece350l_total.xls')
first_sheet = info.sheet_by_index(0)

num_row = first_sheet.nrows
num_col = first_sheet.ncols
lab_std = 6
num_lab = 14
lab_end = lab_std + num_lab - 1
formal = 20
quiz_std = 21
num_quiz = 12
quiz_end = quiz_std + num_quiz - 1

label = first_sheet.row_values(0) # list
usernamelist = first_sheet.col_values(2)
del(usernamelist[0])

record_lab = {}
record_formal = {}
record_quiz = {}
ave = {}
total = {}

# initialize
for name in usernamelist:
	record_lab[name] = []
	record_formal[name] = []
	record_quiz[name] = []
	ave[name] = []
	total[name] = []

from pandas import DataFrame
l1 = usernamelist
l2 = []
l3 = []
l4 = []
l5 = []
l6 = []

# create grade list for lab, formal lab, quiz, average
for idx in range(1,num_row):
	student_info = first_sheet.row_values(idx) 
	name = student_info[2]
	record_lab[name] += student_info[lab_std:lab_end+1]
	record_formal[name] += [student_info[formal]]
	record_quiz[name] += student_info[quiz_std:quiz_end+1]
	# calculate average
	ave[name] += [(sum(record_lab[name]) - min(record_lab[name])) / 13.0]
	ave[name] += record_formal[name]
	ave[name] += [(sum(record_quiz[name]) - min(record_quiz[name])) / 11.0]
	# total
	total[name] = [sum(ave[name])]

	# assign grade
	if total[name][0] >= 97.0:
		total[name] += ['A+']
	elif 93.0 <= total[name][0] < 97.0:
		total[name] += ['A']
	elif 90.0 <= total[name][0] < 93.0:
		total[name] += ['A-']
	elif 87.0 <= total[name][0] < 90.0:
		total[name] += ['B+']
	elif 83.0 <= total[name][0] < 87.0:
		total[name] += ['B']
	elif 80.0 <= total[name][0] < 83.0:
		total[name] += ['B-']
	elif 77.0 <= total[name][0] < 80.0:
		total[name] += ['C+']
	elif 73.0 <= total[name][0] < 77.0:
		total[name] += ['C']
	elif 70.0 <= total[name][0] < 73.0:
		total[name] += ['C-']
	elif 67.0 <= total[name][0] < 70.0:
		total[name] += ['D+']
	elif 60.0 <= total[name][0] < 67.0:
		total[name] += ['D']
	elif total[name][0] < 60.0:
		total[name] += ['F']

	l2.append(ave[name][0])
	l3.append(ave[name][1])
	l4.append(ave[name][2])
	l5.append(total[name][0])
	l6.append(total[name][1])

df = DataFrame({'username':l1, 'ave_lab':l2, 'formal lab':l3, 'ave_quiz':l4, 'total':l5, 'grade':l6})
print df
df.to_excel('ece350_finalgrade.xls', sheet_name='sheet1', index=False)



