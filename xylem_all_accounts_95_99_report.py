import sys
sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *

sql1="""
select distinct
dbr_client,
dbr_cli_ref_no,
dbr_name1,
dbr_no,
dbr_status,
dbr_desk
from CDS.DBR
where dbr_client in ('XYLM95','XYLM96','XYLM97','XYLM98','XYLM99');
"""

# Accepting the filename as an argument or using a variable to write it.
if "filename" not in jm:
	filename = 'Xylem 95-99 All Accounts Report ' + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
	wb = xlsxwriter.Workbook(filename)
else:
	wb = xlsxwriter.Workbook(jm["filename"])


f_text_center = wb.add_format(text_center)
f_text_center.set_bg_color('#afc7ea')
f_text_center.set_bold()
f_text_left_cell = wb.add_format(text_center)
f_text_left = wb.add_format(text_left)

ws = wb.add_worksheet('All Accounts')

ws.write('A1', 'Client ID', f_text_center)
ws.write('B1', 'CLIENT REF #', f_text_center)
ws.write('C1', 'CUSTOMER NAME', f_text_center)
ws.write('D1', 'Debtor Number', f_text_center)
ws.write('E1', 'Account Status', f_text_center)
ws.write('F1', 'Desk ID', f_text_center)
ws.write('G1', 'CRS Note', f_text_center)

ws.set_column(0, 0, 8)
ws.set_column(1, 1, 11)
ws.set_column(2, 2, 16)
ws.set_column(3, 3, 20)
ws.set_column(4, 4, 18)
ws.set_column(5, 5, 18)
ws.set_column(6, 6, 55)

row = 1

for l in sqlSelectList(curs, sql1, ()):
	i = tuple_to_clean_list(l)
	note = return_note('ESC', i[4])
	
	ws.write(row, 0, i[0], f_text_left_cell)
	ws.write(row, 1, i[1], f_text_left_cell)
	ws.write(row, 2, i[2], f_text_left_cell)
	ws.write(row, 3, i[3], f_text_left_cell)
	ws.write(row, 4, i[4], f_text_left_cell)
	ws.write(row, 5, i[5], f_text_left_cell)
	write_note('CRS', i[3], 'summaryNotes', ws, row, 6, f_text_left_cell)
	ws.write(row, 7, '.')

	row += 1
wb.close()
