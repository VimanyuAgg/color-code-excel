# import xlsxwriter
# from xlsxwriter.workbook import Workbook 

# workbook = xlsxwriter.Workbook('consolidated_report.xlsx')
# print workbook
# print "workbook created"

# print workbook.worksheets()
# worksheet= workbook.get_worksheet_by_name("Consolidated_Report__crosstab")
# print worksheet
# print "worksheet created"
# green_format = workbook.add_format()
# green_format.set_pattern(1)
# green_format.set_bg_color('#008000')

# worksheet.write(1,5,green_format)


# worksheet = workbook.add_worksheet('Hello')
# workbook.close()

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
from datetime import datetime, timedelta

wb = openpyxl.load_workbook(filename='consolidated_report.xlsx')
# ws = wb.get_worksheet_by_name('Consolidated_Report__crosstab')

ws = wb['Consolidated_Report__crosstab']

# print ws
# total_entry = ws['A']
# totalRows = 0
# for t in total_entry:
# 	if t.value == None:
# 		break
# 	else:
# 		totalRows +=1

# print totalRows
# print ws.max_row
#print ws['A1'].value
print ws.max_column
greenFill = PatternFill(start_color='10ff00',
                   end_color='10ff00',
                   fill_type='solid')

redFill = PatternFill(start_color='FFFF0000',
				   end_color='FFFF0000',
                   fill_type='solid')

amberFill = PatternFill(start_color='FFC200',
                   end_color='FFC200',
                   fill_type='solid')
blueFill =  PatternFill(start_color='6666ff',
                   end_color='6666ff',
                   fill_type='solid')

wineFill =  PatternFill(start_color='722f37',
                   end_color='722f37',
                   fill_type='solid')

project_status = ws['F']
project_stage = ws['E']
project_per_complete = ws['N']
project_start_date =ws['J']
project_end_date = ws['K']
last_updated  = ws['M']
rag_status = ws['G']
rag_reason = ws['H']
margin_variance = ws['O']
delivery_days_variance = ws['P']
reason_overrun = ws['Q']
delivery_country = ws['T']
project_type= ws['AB']
# to_expiry_date = ws['J']
project_category = ws['AD']
project_revenue = ws['AE']
revenue_source = ws['Z']
amount_invoiced = ws['AF']


for i in range (1,len(project_status)):
	print project_per_complete[i].value > 1
	print type(project_per_complete[i].value)


	# R9C1
	if project_start_date[i].value == None and project_status[i].value != "Booked":
		project_start_date[i].fill = redFill
		project_status[i].fill = redFill

	# R10C1
	if project_end_date[i].value == None and project_status[i].value != "Booked":
		project_end_date[i].fill = redFill
		project_status[i].fill = redFill

	# R10C4
	if (project_end_date[i].value != None and project_start_date[i].value != None) and project_end_date[i].value < project_start_date[i].value:
		project_end_date[i].fill = redFill
		project_start_date[i].fill = redFill

	# R10C2
	if (project_end_date[i].value != None) and (project_end_date[i].value < datetime.now().date() and (project_status[i].value != "Complete" or (project_status[i].value.find("Complete") != -1 and project_status[i].value.find("-") != -1) or project_status[i].value != "Closed")):
		project_end_date[i].fill = redFill
		project_status[i].fill = redFill

	# R10C3
	if (project_end_date[i].value > datetime.now().date() and (project_status[i].value == "Complete" or project_status[i].value == "Closed" or (project_status[i].value.find("Complete") != -1 and project_status[i].value.find("-") != -1))):
		project_status[i].fill = redFill
		project_end_date[i].fill = redFill


	#R11C1
	if (last_updated[i].value == None and project_status[i].value != "Booked"):
		last_updated[i].fill = redFill
		project_status[i].fill = "Booked"

	# R11C2
	if (last_updated[i].value != None) and (last_updated[i].value < datetime.now().date() - timedelta(days=14)):
		last_updated[i].fill = redFill

	# R11C3
	if (last_updated[i].value != None) and (last_updated[i].value > datetime.now().date()):
		last_updated[i].value = redFill

	# R12
	if(rag_status[i].value == None):
		rag_status[i].fill = redFill


	#R13
	## RAG REASON 404

	#R14
	if (margin_variance[i].value > -0.05):
		margin_variance[i].fill = redFill

	#R15
	if (delivery_days_variance[i].value > -0.1):
		delivery_days_variance[i].fill = redFill

	#R16
	if (reason_overrun[i].value == None):
		reason_overrun[i].fill = redFill

	#R23
	if (delivery_country[i].value == None):
		delivery_country[i].fill = redFill

	#R24
	if (project_type[i].value == None):
		project_type[i].fill = redFill

	# # R25
	# if (to_expiry_date[i].value == None):
	# 	to_expiry_date[i].fill = redFill

	# R26
	if (project_category[i].value == None):
		project_category[i].fill = redFill

	# R27C1
	if (project_revenue[i].value == None):
		project_revenue[i].fill = redFill

	# R27 C2
	if (project_revenue[i].value != None and (revenue_source[i].value.find("PO") != -1 )):
		project_revenue[i].fill = redFill
		# revenue_source[i].fill = redFill

	# R27C3
	if (project_revenue[i].value != 0 and (revenue_source[i].value == None or revenue_source[i].value == "CSAT" or revenue_source[i].value == "FOC" or revenue_source[i].value == "Reclass")):
		project_revenue[i].fill = redFill


	# R28
	if (amount_invoiced[i].value != None and project_revenue[i].value != None) and (amount_invoiced[i].value > project_revenue[i].value):
		amount_invoiced[i] .fill = redFill

	# R8 case 1
	if project_per_complete[i].value > 1:
		project_per_complete[i].fill = redFill
	else: #R8C3
		if (project_status[i].value == "Complete" or 
		   (project_status[i].value.find("Complete") != -1 and project_status[i].value.find("-") != -1) or (project_status[i].value == "Closed")):
			project_status[i].fill = redFill
			project_per_complete[i].fill = redFill
			
	#Req R7 case 1 & R8C2:
	if project_per_complete[i].value != None and ((int(project_per_complete[i].value) >= 1) and (project_status[i].value != "Complete" 
		or project_status[i].value !="Closed")):
		project_status[i].fill = redFill
		project_per_complete[i].fill = redFill
		continue

	#Req R7 case 2:
	if (project_status[i].value.find("In Progress") !=-1 or project_stage[i].value.find("On Hold") !=-1 or project_status[i].value.find("Complete") !=1) and (project_status[i].value.find("-") !=-1):
		if project_stage[i].value != "Opportunity":
			project_stage[i].fill = redFill
			project_status[i].fill = redFill
		continue #No more rules for WAR

	# R7 case 3:
	if (project_status[i].value == "In Progress" or project_status[i].value == "Booked" 
		or project_status[i].value == "On Hold" or project_status[i].value == "Complete") and project_stage[i].value != "Awarded":
		project_stage[i].fill = redFill
		project_status[i].fill = redFill


# for i in range(1,len(project_status)):
	









# for c in project_status:
# 	#print type(c.value)
# 	if c.value == "On Hold":
# 		c.fill = redFill
# 	elif c.value == "In Progress":
# 		c.fill = amberFill
# 	elif c.value =="Complete":
# 		c.fill = greenFill
# 	elif c.value == "Booked":
# 		c.fill = blueFill


# for col in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
# 	for cell in col:
# 		if cell.value == None:
# 			cell.fill = wineFill


wb.save('updated.xlsx')