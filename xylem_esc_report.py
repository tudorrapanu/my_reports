import sys

sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
import datetime

sql1 = """
select
dbr_client,
dbr_cli_ref_no,
dbr_name1,
DATEFORMAT(dbr_assign_date_o, 'YYYY-MM-DD'),
dbr_no,
dbr_cl_misc_1,
dbr_cl_misc_2
from cds.dbr
inner join cds.dat on dat_Dbr_no = dbr_no
where dbr_client in ('XYLM95','XYLM96','XYLM97')
and dat_action_code = 'ESC'
and dbr_close_date_o is null
and dat_trx_date_o between current date - 8 and current date - 1
group by dbr_client,
dbr_cli_ref_no,
dbr_name1,
dbr_assign_date_o,
dbr_no,
dat_trx_date_o,
dbr_cl_misc_1,
dbr_cl_misc_2;
"""

if "filename" not in jm:
    filename = 'xylem_esc_report' + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
    wb = xlsxwriter.Workbook(filename)
else:
    wb = xlsxwriter.Workbook(jm["filename"])

f_text_center = wb.add_format(text_center)
f_text_center.set_bg_color('#afc7ea')
f_text_center.set_bold()
f_text_left_cell = wb.add_format(text_center)
f_text_left = wb.add_format(text_left)

ws = wb.add_worksheet('Escalation')

ws.write('A1', 'Client ID', f_text_center)
ws.write('B1', 'CLIENT REF #', f_text_center)
ws.write('C1', 'Invoice Numbers', f_text_center)
ws.write('D1', 'CUSTOMER NAME', f_text_center)
ws.write('E1', 'Placement Date', f_text_center)
ws.write('F1', 'Debtor Misc 1', f_text_center)
ws.write('G1', 'Debtor Misc 2', f_text_center)
ws.write('H1', 'Escalation Reason', f_text_center)

ws.set_column(0, 0, 8)
ws.set_column(1, 1, 11)
ws.set_column(2, 2, 16)
ws.set_column(3, 3, 20)
ws.set_column(4, 4, 18)
ws.set_column(5, 5, 18)
ws.set_column(6, 6, 18)
ws.set_column(7, 7, 55)

row = 1

for l in sqlSelectList(curs, sql1, ()):
    i = tuple_to_clean_list(l)
    note = return_note('ESC', i[4])

    ws.write(row, 0, i[0], f_text_left_cell)
    ws.write(row, 1, i[1], f_text_left_cell)
    ws.write(row, 2, note['widgetINVList'], f_text_left_cell)
    ws.write(row, 3, i[2], f_text_left_cell)
    ws.write(row, 4, i[3], f_text_left_cell)
    ws.write(row, 5, i[5], f_text_left_cell)
    ws.write(row, 6, i[6], f_text_left_cell)
    ws.write(row, 7, note['escalationReason'], f_text_left_cell)
    ws.write(row, 8, '.')

    row += 1
wb.close()
