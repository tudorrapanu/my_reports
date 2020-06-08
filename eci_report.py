import sys

sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
import datetime

sql1 = """
select 
dbr_client,
dbr_no,
dbr_name1,
dbr_cli_ref_no,
(select ifnull(sum(ivt_paid), 0)
        from CDS.IVT
        where ivt_dbr_no = dbr_no) as amt_paid,
(select ifnull(sum(ivt_amount-ivt_paid), 0) 
        from CDS.IVT 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=0) as bucket0, 
(select ifnull(sum(ivt_amount-ivt_paid), 0) 
        from CDS.IVT 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=30) as bucket30,
(select ifnull(sum(ivt_amount-ivt_paid), 0)  
        from CDS.IVT 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=60) as bucket60, 
(select ifnull(sum(ivt_amount-ivt_paid), 0)  
        from CDS.IVT 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=90) as bucket90, 
(select ifnull(sum(ivt_amount-ivt_paid), 0) 
        from CDS.IVT
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=120) as bucket120,
(select ifnull(sum(ivt_amount-ivt_paid), 0) 
        from CDS.IVT 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=180) as bucket180,
(select ifnull(sum(ivt_amount-ivt_paid), 0)  
        from CDS.IVT 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=181) as bucket181,  
dbr_status, 
dbr_last_worked_o,
(select max(dat_Seq_no) 
        from CDS.DAT 
        where dat_dbr_no = dbr_no 
        and dat_note = 'Y') as max_seq,
(select GROUP_CONCAT(DNT_NOTE ORDER BY DNT_SEQ_SUB SEPARATOR  '') 
        from CDS.DNT 
        where dnt_dbr_no = dbr_no 
        and dnt_seq_no = max_seq) as last_note,
dbr_desk
from CDS.DBR
where dbr_client in ('ECI099', 'ECI098', 'ECI097')
group by
dbr_status, 
dbr_last_worked_o,
dbr_client,
dbr_no,
dbr_name1,
dbr_cli_ref_no,
dbr_desk
order by dbr_client;
"""

if "filename" not in jm:
    filename = 'eci_report' + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
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

ws = wb.add_worksheet('ECI Report')

ws.merge_range('A1:B6', '')
ws.insert_image('A1', 'logo.png', {'x_scale': 0.35, 'y_scale': 0.35})

ws.merge_range('C1:P3', '')
ws.merge_range('C4:P6', '')

ws.write('C1', 'ECI Accounts Report', f_text_header)
ws.write('C4', run_date.strftime('%Y-%m-%d'), f_text_header)

ws.write('A7', 'Client ID', f_text_center)
ws.write('B7', 'Debtor Number', f_text_center)
ws.write('C7', 'Name', f_text_center)
ws.write('D7', 'Client Reference Number', f_text_center)
ws.write('E7', 'Amount Received', f_text_center)
ws.write('F7', '<0 Days', f_text_center)
ws.write('G7', '0-30 Days', f_text_center)
ws.write('H7', '31-60 Days', f_text_center)
ws.write('I7', '61-90 Days', f_text_center)
ws.write('J7', '91-120 Days', f_text_center)
ws.write('K7', '121-180 Days', f_text_center)
ws.write('L7', '181+ Days', f_text_center)
ws.write('M7', 'Account Status', f_text_center)
ws.write('N7', 'Last Worked Date', f_text_center)
ws.write('O7', 'Last Note', f_text_center)
ws.write('P7', 'Desk ID', f_text_center)

ws.set_column(0, 0, 10)
ws.set_column(1, 1, 15)
ws.set_column(2, 3, 30)
ws.set_column(4, 10, 10)
ws.set_column(11, 13, 15)
ws.set_column(14, 14, 30)
ws.set_column(15, 15, 8)

row = 7

for l in sqlSelectList(curs, sql1, ()):
    i = tuple_to_clean_list(l)

    ws.write(row, 0, i[0], f_text_normal)     # dbr_client
    ws.write(row, 1, i[1], f_text_normal)     # dbr_no
    ws.write(row, 2, i[2], f_text_normal)     # dbr_name1
    ws.write(row, 3, i[3], f_text_normal)     # dbr_cli_ref_no
    ws.write(row, 4, i[4], f_text_normal)     # amt_paid
    ws.write(row, 5, i[5], f_text_normal)     # bucket0
    ws.write(row, 6, i[6], f_text_normal)     # bucket30
    ws.write(row, 7, i[7], f_text_normal)     # bucket60
    ws.write(row, 8, i[8], f_text_normal)     # bucket90
    ws.write(row, 9, i[9], f_text_normal)     # bucket120
    ws.write(row, 10, i[10], f_text_normal)   # bucket180
    ws.write(row, 11, i[11], f_text_normal)   # bucket181
    ws.write(row, 12, i[12], f_text_normal)   # dbr_status
    dt = i[13].strftime('%Y-%m-%d')
    ws.write(row, 13, dt, f_text_normal)      # dbr_last_worked_o
    ws.write(row, 14, i[15], f_text_left)     # last_note
    ws.write(row, 15, i[16], f_text_normal)   # dbr_desk
    ws.write(row, 16, '.')

    row += 1
wb.close()
