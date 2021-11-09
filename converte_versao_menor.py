#!/usr/bin/python3

import re, os, sys
import xlsxwriter
sep="|"

fp = open(sys.argv[1], "r")

workbook = xlsxwriter.Workbook(sys.argv[2])
ws = workbook.add_worksheet('Data') 

# Antes: lin0=["Pront.", "Reg.", "Nome do paciente", "Idade", "Sexo", "Cidade", "Data", "Horas", "C", "Medico Id", "Medico Nome", "Convenio ID", "Convenio"]
# Agora: lin0=["Cnv", "Reg.", "Nome do paciente", "Hora", "Pront.", "Especialidade", "Medico", "Convenio", "R"]

lin0=["Data", "Cnv", "Reg.", "Nome do paciente", "Hora", "Pront.", "Especialidade", "Medico", "Convenio"]

for i, l in enumerate(lin0):
	ws.write(0, i, l)
	
lin = 1
data = "null"
for l in fp.readlines():	#
	if 'Data do atendimento:' in l:
		data = str(re.search('\d{2}\/\d{2}\/\d{4}', l.strip())[0]).strip()
		
	elif re.search('\d+\s\d+\s\w+', l):		
		reg=[]

		# Col 1 Cnv
		#reg.append('(\d{3}\/\d{2})(?=\s\d+\s.+)')
		reg.append('(\d+)(?=\s\d+\s)')
						
		# Col 2 Reg.
		reg.append('(?<=\d\s)(\d{8})(?=\s\w+)')
		
		# Col 3 Nome do paciente
		reg.append('(?<=\s\d{8}\s)(.*)(?=\s\d{2}:\d{2}\s)')
		
		# Col 4 Hora
		reg.append('(?<=\s)(\d{2}:\d{2})(?=\s\d{3}\/\d{2}\s)')
		
		# Col 5 Pront
		reg.append('(?<=\d{2}:\d{2}\s)(\d{3}\/\d{2})(?=\s\w+)')
		
		# Col 6 Especialidade
		reg.append('(?<=\d{3}\/\d{2}\s)([A-Z].*[a-z]\s)')
		
		# Col 7 Medico
		reg.append('(?<=[a-z]\s)([A-Z]{2}.*)(?= UNIMED| SUS| DOAÇÃO| PARTICULAR| COTA SMS| CASSI| SERPRAM| CLIMEPE| POLICIA MILITAR)')
		
		# Col 8 Convenio
		reg.append('((UNIMED | SUS | DOAÇÃO | PARTICULAR | COTA SMS| CASSI | SERPRAM | CLIMEPE | POLICIA MILITAR).*)(?=\|)')

		ws.write(lin, 0, data)				
		for i, r in enumerate(reg):	
				
			try:
				o = str(re.search(r, l.strip())[0]).strip() 
				print(o + str(sep), end='')
				ws.write(lin, i+1, o)
							
			except:
				print("#################################################################################")
				print("#############Erro -------------> " + str(l.strip()))
				print("Regex:" + str(r))
				
		#print()
		lin+=1
		
fp.close()
workbook.close()
