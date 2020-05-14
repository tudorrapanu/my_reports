import sys
sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
import datetime

sql1 = """
select
dbr_no,
dbr_cli_ref_no,
dbr_name1,
dateformat(dbr_assign_date_o,"YYYY-MM-DD"),
count (*) as no_inv,
sum(ivt_amount) as total_amt,
sum(ivt_paid) as paid_amt
from cds.dbr
inner join cds.ivt on dbr_no = ivt_dbr_no
where 
dbr_client in ('GEHC51')
and dateformat(dbr_assign_date_o,"YYYY-MM-DD") > '2016-12-31'
and dbr_assign_amt>1000
and dateformat(ivt_due_date,"YYYY-MM-DD") > '2016-12-31'
group by
dbr_no,
dbr_cli_ref_no,
dbr_name1,
dbr_assign_date_o
having count (*) > 1;
"""

if "filename" not in jm:
	filename = 'ge_inv_report' + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
	wb = xlsxwriter.Workbook(filename)
else:
	wb = xlsxwriter.Workbook(jm["filename"])

f_text_center = wb.add_format(text_center)
f_text_center.set_bg_color('#afc7ea')
f_text_center.set_bold()
f_text_left_cell = wb.add_format(text_center)
f_text_left = wb.add_format(text_left)

ws = wb.add_worksheet('Accounts')

ws.write('A1', 'Debtor Number', f_text_center)
ws.write('B1', 'Client Reference', f_text_center)
ws.write('C1', 'Debtor Name', f_text_center)
ws.write('D1', 'Placement Date', f_text_center)
ws.write('E1', 'Number Of Due Invoices This Year', f_text_center)
ws.write('F1', 'Total Amount Of The Due Invoices', f_text_center)
ws.write('G1', 'Total Received On Those Due Invocies', f_text_center)

ws.set_column(0, 0, 8)
ws.set_column(1, 1, 11)
ws.set_column(2, 2, 20)
ws.set_column(3, 3, 16)
ws.set_column(4, 4, 18)
ws.set_column(5, 5, 18)
ws.set_column(6, 6, 18)

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
	ws.write(row, 6, i[6], f_text_left_cell)

	row += 1
wb.close()
