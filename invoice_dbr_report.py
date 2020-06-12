import sys

sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
import datetime

sql1 = """
select
dbr_client,
dbr_no,
dbr_cli_ref_no,
dbr_assign_amt,
dbr_assign_date_o,
ivt_ivt_no,
ivt_ivt_date,
ivt_due_date,
ivt_amount,
ivt_amount-ivt_paid as ivt_due
from CDS.DBR
inner join CDS.IVT on ivt_dbr_no = dbr_no
where dbr_no = %s;
"""

dbr_no = jm[debtor_number]

if "filename" not in jm:
    filename = 'invoice_report_' + dbr_no + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
    wb = xlsxwriter.Workbook(filename)
else:
    wb = xlsxwriter.Workbook(jm["filename"])

f_text_center = wb.add_format(text_center)
f_text_center.set_bg_color('#afc7ea')
f_text_center.set_bold()
f_text_normal = wb.add_format(text_center)
f_text_header = wb.add_format(text_center)
f_text_header.set_bold()
f_text_header.set_font_color('navy')
f_text_header.set_font_size(18)
f_text_left_cell = wb.add_format(text_center)
f_text_left = wb.add_format(text_left)
date_format = wb.add_format({'num_format': '%Y-%m-%d', 'align': 'center'})

ws = wb.add_worksheet('Invoices Report')

ws.merge_range('A1:B6', '')
ws.insert_image('A1', 'logo.png', {'x_scale': 0.35, 'y_scale': 0.35})

ws.merge_range('C1:Q3', '')
ws.merge_range('C4:Q6', '')

ws.write('C1', 'Invoice Details for Debtor Number Report', f_text_header)
ws.write('C4', run_date.strftime('%Y-%m-%d'), f_text_header)

ws.write('A7', 'Client ID', f_text_center)
ws.write('B7', 'Debtor Number', f_text_center)
ws.write('C7', 'Client Reference Number', f_text_center)
ws.write('D7', 'Assigned Amount', f_text_center)
ws.write('E7', 'Assigned Date', f_text_center)
ws.write('F7', 'Invoice Number', f_text_center)
ws.write('G7', 'Invoice Date', f_text_center)
ws.write('H7', 'Invoice Due Date', f_text_center)
ws.write('I7', 'Invoice Amount', f_text_center)
ws.write('J7', 'Amount Due', f_text_center)

ws.set_column(0, 0, 10)
ws.set_column(1, 2, 15)
ws.set_column(3, 4, 30)
ws.set_column(5, 5, 15)
ws.set_column(6, 9, 10)

row = 7

for l in sqlSelectList(curs, sql1, (dbr_no)):
    i = tuple_to_clean_list(l)

    ws.write(row, 0, i[0], f_text_normal)     # dbr_client
    ws.write(row, 1, i[1], f_text_normal)     # dbr_no
    ws.write(row, 2, i[2], f_text_normal)     # dbr_cli_ref_no
    ws.write(row, 3, i[3], f_text_normal)     # dbr_assign_amt
    dt = i[4].strftime('%Y-%m-%d')
    ws.write(row, 4, dt, f_text_normal)       # dbr_assign_date_o
    ws.write(row, 5, i[5], f_text_normal)     # ivt_ivt_no
    dt = i[6].strftime('%Y-%m-%d')
    ws.write(row, 6, dt, f_text_normal)       # ivt_ivt_date
    dt = i[7].strftime('%Y-%m-%d')
    ws.write(row, 7, dt, f_text_normal)       # ivt_due_date
    ws.write(row, 8, i[8], f_text_normal)     # ivt_amount
    ws.write(row, 9, i[9], f_text_normal)     # ivt_due
    ws.write(row, 10, '.')

    row += 1
wb.close()

