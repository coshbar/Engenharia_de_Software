import gspread
import math
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('EngenhariaSoftwareDesafio-eb43a06052d0.json', scope)
client = gspread.authorize(credentials)
sheet = client.open("Engenharia de Software - Desafio Gabriel Klimczak").sheet1

#getting all data into a list
sheet_list = sheet.get_all_values()

#get the total number of lectures in the semester
lecture_number = sheet_list[1]
lecture_number = int(''.join(filter(str.isdigit, lecture_number[0])))

#determinates where the student data starts
start_student = 3

#columns for specific data
nonattendance_location = 2
p1_location = 3
p2_location = 4
p3_location = 5
situation_location = 7
naf_location = 8

#row for cell update
row = 4

for student in sheet_list[start_student::]:
	#checks if student has less than 75% attendance
	if int(student[nonattendance_location]) > lecture_number * 0.25:
		#do stuff: Reprovado por Falta
		sheet.update_cell(row, situation_location, 'Reprovado por Falta')
		sheet.update_cell(row, naf_location, '0')
		
		#cmd log
		print('Reprovado por Falta')

		#move to next row
		row += 1
	else:
		m = (int(student[p1_location]) + int(student[p2_location]) + int(student[p3_location])) / 3
		
		if m < 50:
			#do stuff: Reprovado por Nota
			sheet.update_cell(row, situation_location, 'Reprovado por Nota')
			sheet.update_cell(row, naf_location, '0')
		
			#cmd log
			print('Reprovado por Nota')
		elif 50 <= m < 70:
			#do stuff: Exame Final
			naf = 50 * 2 - m
			sheet.update_cell(row, situation_location, 'Exame Final')
			sheet.update_cell(row, naf_location, math.ceil(naf))
		
			#cmd log
			print(math.ceil(naf))
		else:
			#do stuff: Aprovado
			sheet.update_cell(row, situation_location, 'Aprovado')
			sheet.update_cell(row, naf_location, '0')
					
			#cmd log
			print('Aprovado')

		#move to next row
		row += 1

			
		
