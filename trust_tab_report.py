import sys

sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
import datetime

sql1 = """
select
date_format(trs_trx_date_o, '%Y-%m-%d'),
TRS_USERID,
trs_note_1,
trs_trust_code,
trs_stmt_byte,
trs_mop,
trs_desk,
trs_amt,
trs_comm_amt,
trs_ar_agency,
trs_note_2,
remit_no,
date_format(remit_date_o, '%Y-%m-%d')
from CDS.TRS
inner join DBA.trs_remit on trs_remit.remit_id = remit_no
where trs_dbr_no = '5175264'
"""

if "filename" not in jm:
    filename = 'trust_tab_report' + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
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

ws = wb.add_worksheet('Trust Report')

ws.merge_range('A1:B6', '')
ws.insert_image('A1', 'logo.png', {'x_scale': 0.35, 'y_scale': 0.35})

ws.merge_range('C1:L3', '')
ws.merge_range('C4:L6', '')

ws.write('C1', 'ECI Accounts Report', f_text_header)
ws.write('C4', run_date.strftime('%Y-%m-%d'), f_text_header)

ws.write('A7', 'Post Date', f_text_center)
ws.write('B7', 'User', f_text_center)
ws.write('C7', 'Invoice', f_text_center)
ws.write('D7', 'Code', f_text_center)
ws.write('E7', 'Type', f_text_center)
ws.write('F7', 'MOP', f_text_center)
ws.write('G7', 'Desk', f_text_center)
ws.write('H7', 'Amount', f_text_center)
ws.write('I7', 'Fee', f_text_center)
ws.write('J7', 'Currency', f_text_center)
ws.write('K7', 'FX Rate', f_text_center)
ws.write('L7', 'Remit', f_text_center)
ws.write('M7', 'Rent Date', f_text_center)

ws.set_column(0, 2, 10)
ws.set_column(3, 3, 8)
ws.set_column(4, 4, 15)
ws.set_column(5, 6, 10)
ws.set_column(7, 8, 15)
ws.set_column(9, 11, 10)
ws.set_column(12, 12, 15)

row = 7

for l in sqlSelectList(curs, sql1, ()):
    i = tuple_to_clean_list(l)

    ws.write(row, 0, i[0], f_text_normal)     # trs_trx_date_o
    ws.write(row, 1, i[1], f_text_normal)     # TRS_USERID
    ws.write(row, 2, i[2], f_text_normal)     # trs_note_1
    ws.write(row, 3, i[3], f_text_normal)     # trs_trust_code
    if (i[4]=='A'):
        ws.write(row, 4, 'Agency', f_text_normal) # trs_stmt_byte
    else if(i[4]=='C'):
        ws.write(row, 4, 'Client', f_text_normal)
    else:
        ws.write(row, 4, i[4], f_text_normal)
    ws.write(row, 5, i[5], f_text_normal)     # trs_mop
    ws.write(row, 6, i[6], f_text_normal)     # trs_desk
    ws.write(row, 7, i[7], f_text_normal)     # trs_amt
    ws.write(row, 8, i[8], f_text_normal)     # trs_comm_amt
    ws.write(row, 9, i[9], f_text_normal)     # trs_ar_agency
    ws.write(row, 10, i[10], f_text_normal)   # trs_note_2
    ws.write(row, 11, i[11], f_text_normal)   # remit_no
    ws.write(row, 12, i[12], f_text_normal)   # remit_date_o
    ws.write(row, 13, '.')

    row += 1
wb.close()

