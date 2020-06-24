import sys

sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
import datetime

sql1 = """
select
trs_note_1 as Invoice,
date_format(trs_trx_date_o, '%Y-%m-%d') as Post_Date,
date_format(ivt_due_date, '%Y-%m-%d') as Due_Date,
TRS_USERID as User_,
trs_trust_code as Code,
trs_stmt_byte as Type,
trs_mop as MOP,
trs_desk as Desk,
trs_amt as Amount,
trs_comm_amt as Fee,
trs_ar_agency as Currency
from CDS.TRS
inner join CDS.IVT on trs_note_1 = ivt_ivt_no
where trs_dbr_no = '6420578'
and trs_trx_date_o < ivt_due_date
"""

if "filename" not in jm:
    filename = 'inv_paid_before_due_6420578' + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
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

ws = wb.add_worksheet('Invoices Report')

ws.merge_range('A1:B6', '')
ws.insert_image('A1', 'logo.png', {'x_scale': 0.35, 'y_scale': 0.35})

ws.merge_range('C1:K3', '')
ws.merge_range('C4:K6', '')

ws.write('C1', 'Invoices Paid Before Due Date for 6420578', f_text_header)
ws.write('C4', run_date.strftime('%Y-%m-%d'), f_text_header)

ws.write('A7', 'Invoice', f_text_center)
ws.write('B7', 'Post Date', f_text_center)
ws.write('C7', 'Due Date', f_text_center)
ws.write('D7', 'User', f_text_center)
ws.write('E7', 'Code', f_text_center)
ws.write('F7', 'Type', f_text_center)
ws.write('G7', 'MOP', f_text_center)
ws.write('H7', 'Desk', f_text_center)
ws.write('I7', 'Amount', f_text_center)
ws.write('J7', 'Fee', f_text_center)
ws.write('K7', 'Currency', f_text_center)

ws.set_column(0, 0, 10)
ws.set_column(1, 2, 15)
ws.set_column(3, 4, 30)
ws.set_column(5, 5, 15)
ws.set_column(6, 10, 10)

row = 7

for l in sqlSelectList(curs, sql1, ()):
    i = tuple_to_clean_list(l)

    ws.write(row, 0, i[0], f_text_normal)     # trs_note_1
    ws.write(row, 1, i[1], f_text_normal)     # trs_trx_date_o
    ws.write(row, 2, i[2], f_text_normal)     # ivt_due_date
    ws.write(row, 3, i[3], f_text_normal)     # trs_userid
    ws.write(row, 4, i[4], f_text_normal)     # trs_trust_code
    ws.write(row, 5, i[5], f_text_normal)     # trs_stmt_byte
    ws.write(row, 6, i[6], f_text_normal)     # trs_mop
    ws.write(row, 7, i[7], f_text_normal)     # trs_desk
    ws.write(row, 8, i[8], f_text_normal)     # trs_amt
    ws.write(row, 9, i[9], f_text_normal)     # trs_comm_amt
    ws.write(row, 10, i[10], f_text_normal)   # trs_ar_agency
    ws.write(row, 11, '.')

    row += 1
wb.close()


