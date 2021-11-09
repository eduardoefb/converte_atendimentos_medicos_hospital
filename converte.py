#!/usr/bin/python3

import re, os, sys
import xlsxwriter
sep="|"

fp = open(sys.argv[1], "r")

workbook = xlsxwriter.Workbook(sys.argv[2])
ws = workbook.add_worksheet('Data') 

lin0=["Pront.", "Reg.", "Nome do paciente", "Idate", "Sexo", "Cidade", "Data", "Horas", "C", "Medico Id", "Medico Nome", "Convenio ID", "Convenio"]

for i, l in enumerate(lin0):
	ws.write(0, i, l)
	
lin = 1
for l in fp.readlines():	#
	if re.search('\d+\/\d+\s\d+\s\w+', l):
		reg=[]
		# Col 1 Pront.
		reg.append('(\d{3}\/\d{2})(?=\s\d+\s.+)')
				
		# Col 2 Reg.
		reg.append('(?<=\d{3}\/\d{2}\s)(\d+)(?=\s\w+\s)')
		
		# Col 3 Nome do paciente
		#reg.append('(?<=\d{3}\/\d{2}\s\d{5}\s)(.*)(?=\s+\d{3}\s(F|M)\w+\s\w+\s)')
		reg.append('(?<=\d{5}\s)(.*)(?=\s+\d{3}\s(F|M)\w+\s\w+\s)')
		
		# Col 4 Idade
		reg.append('(?<=\s)(\d+)(?=\s(F|M)\w{3})')
		
		# Col 5 Sexo
		reg.append('(?<=\d{3}\s)((F|M)\w{3})(?=\s\w+)')
		
		# Col 6 Cidade
		reg.append('(?<=\s\d{3}\s(F|M)\w{3})(.*)(?=\s+\d{2}\/\d{2}\/\d{4}\s)')
		
		# Col 7 Data
		reg.append('(?<=\s)(\d{2}\/\d{2}\/\d{4})(?=\s\d{2}:\d{2}\s)')
		
		# Col 8 Hora
		reg.append('(?<=\s\d{2}\/\d{2}\/\d{4}\s)(\d{2}:\d{2})(?=\s\w\s\d{5})')
		
		# Col 9 T
		reg.append('(?<=\d{4}\s\d{2}:\d{2}\s)(\w)(?=\s\d{6})')
		
		# Col 10 Medico Codigo
		reg.append('(?<=\d{2}:\d{2}\s\w\s)(\d+)(?=\s\w+)')
		
		# Col 11 Medico Nome
		reg.append('(?<=\d{6}\s)(.*)(?=\s\d{3}\s\w{3}\s)')
		
		# Col 12 Convenio Codigo
		reg.append('(?<=\s)(\d{3})(?=\s\w{3}\s)')
		
		# Col 13 Convenio Nome
		reg.append('(?<=\s\d{3}\s)(\w{3})(?=\s)')
		
		
		for i, r in enumerate(reg):	
				
			try:
				o = str(re.search(r, l.strip())[0]).strip() 
				#print(o + str(sep), end='')
				ws.write(lin, i, o)

				
				
			except:
				print("#################################################################################")
				print("#############Erro -------------> " + str(l.strip()))
				print("Regex:" + str(r))
				
		#print()
		lin+=1
		
fp.close()
workbook.close()
