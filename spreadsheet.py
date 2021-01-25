import gspread
import math
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('EngenhariaSoftwareDesafio-eb43a06052d0.json', scope)
client = gspread.authorize(credentials)
sheet = client.open("Engenharia de Software - Desafio Gabriel Klimczak").sheet1

#getting all data into a list
sheet_list = sheet.get_all_values()

#cell location
cell_id = sheet.find("Matricula")
cell_nonattendance = sheet.find("Faltas")
cell_p1 = sheet.find("P1")
cell_p2 = sheet.find("P2")
cell_p3 = sheet.find("P3")
cell_situation = sheet.find("Situação")
cell_naf = sheet.find("Nota para Aprovação Final")

#get the total number of lectures in the semester
lecture_number = sheet_list[1]
lecture_number = int(''.join(filter(str.isdigit, lecture_number[0])))

#row for cell update
row = cell_id.row + 1

for student in sheet_list[cell_id.row::]:
	#checks if student has less than 75% attendance
	if int(student[cell_nonattendance.col - 1]) > lecture_number * 0.25:
		#do stuff: Reprovado por Falta
		sheet.update_cell(row, cell_situation.col, 'Reprovado por Falta')
		sheet.update_cell(row, cell_naf.col, '0')
		
		#cmd log
		print('Reprovado por Falta')

		#move to next row
		row += 1
	else:
		m = (int(student[cell_p1.col - 1]) + int(student[cell_p2.col - 1]) + int(student[cell_p3.col - 1])) / 3
		if m < 50:
			#do stuff: Reprovado por Nota
			sheet.update_cell(row, cell_situation.col, 'Reprovado por Nota')
			sheet.update_cell(row, cell_naf.col, '0')
		
			#cmd log
			print('Reprovado por Nota')
		elif 50 <= m < 70:
			#do stuff: Exame Final
			naf = 50 * 2 - m
			sheet.update_cell(row, cell_situation.col, 'Exame Final')
			sheet.update_cell(row, cell_naf.col, math.ceil(naf))
		
			#cmd log
			print('Exame Final', math.ceil(naf))
		else:
			#do stuff: Aprovado
			sheet.update_cell(row, cell_situation.col, 'Aprovado')
			sheet.update_cell(row, cell_naf.col, '0')
					
			#cmd log
			print('Aprovado')

		#move to next row
		row += 1			
		
