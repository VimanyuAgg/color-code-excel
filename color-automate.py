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

project_status = ws['E']



for c in project_status:
	#print type(c.value)
	if c.value == "On Hold":
		c.fill = redFill
	elif c.value == "In Progress":
		c.fill = amberFill
	elif c.value =="Complete":
		c.fill = greenFill
	elif c.value == "Booked":
		c.fill = blueFill


for col in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
	for cell in col:
		if cell.value == None:
			cell.fill = wineFill


wb.save('updated.xlsx')