import xlrd
import xlwt
rb = xlrd.open_workbook('e:/list.xls',formatting_info=True)
sheet = rb.sheet_by_index(0)
		
font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.colour_index = 2
font0.bold = True

week = 1
wb = xlwt.Workbook()

while week < 5:
	ws = wb.add_sheet(str(week) + ' week')
	j = 0
	for rownum in range(sheet.nrows):
		row = sheet.row_values(rownum)
		if str(week) in row[1] or row[1] == '':
			i = 0
			for content in row: 
				ws.write(j, i, content)
				i += 1
			j += 1
	week += 1
			
wb.save('e:/Shedule.xls')