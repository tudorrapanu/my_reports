import sys

sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
import datetime

############ QUERRIES ############

sql1 = """
SELECT clt_no,
MIN(POST_DATE),
MAX(POST_DATE),
(SELECT remit_currency FROM DBA.trs_remit WHERE remit_id = '%s') AS currency
FROM DBA.trs_archive
WHERE remit_id = '%s'
GROUP BY clt_no;
"""

sql2 = """
SELECT
DBR.DBR_CLI_REF_NO,
DBR.DBR_NAME1,
IVT_IVT_NO,
TRS_TRX_DATE_O,
TRS_AMT,
TRS_COMM_AMT,
REMIT_TO,
DUE_FROM,
dbr_assign_amt - dbr_recvd_tot as total_due,
TRS_STMT_BYTE as type,
TRS_TRUST_CODE
from CDS.TRS
inner join CDS.IVT on trs_note_1 = ivt_ivt_no
inner join CDS.DBR on trs_dbr_no = DBR.dbr_no
inner join DBA.trs_archive on trs_dbr_no = trs_archive.dbr_no
WHERE dbr_client = '%s'
AND TRS_STMT_BYTE = '%s'
AND remit_id = '%s';
"""

############ PARAMETERS ############
remit_id = jm["remit_id"]

for l in sqlSelectList(curs, sql1, (remit_id, remit_id)):
	i = tuple_to_clean_list(l)
	client_id = i[0]
	min_trust_date = i[1]
	max_trust_date = i[2]
	currency = i[3]

	if "filename" not in jm:
	    filename = remit_id + '-' + client_id + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
	    wb = xlsxwriter.Workbook(filename)
	else:
	    wb = xlsxwriter.Workbook(jm["filename"])

	############ FORMATS ############

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

	############# MAPPING ############

	ws = wb.add_worksheet('Statement')

	ws.merge_range('A1:B6', '')
	ws.insert_image('A1', 'logo.png', {'x_scale': 0.35, 'y_scale': 0.35})

	ws.merge_range('C2:K2', '')
	ws.merge_range('C3:K3', '')

	ws.write('C2', client_id, f_text_header)
	ws.write('C3', 'D&S Statement for period ' + min_trust_date.strftime('%Y-%m-%d') + ' through ' + max_trust_date.strftime('%Y-%m-%d'), f_text_header)

	ws.merge_range('A7:K7', '')
	ws.write('A7', 'Paid To Agency Detail', f_text_normal)
	ws.write('A8', 'Account Number', f_text_center)
	ws.write('B8', 'Debtor Number', f_text_center)
	ws.write('C8', 'Payment Memo', f_text_center)
	ws.write('D8', 'Post Date', f_text_center)
	ws.write('E8', 'Principal', f_text_center)
	ws.write('F8', 'Commission', f_text_center)
	ws.write('G8', 'Remit To', f_text_center)
	ws.write('H8', 'Due From', f_text_center)
	ws.write('I8', 'Total Due', f_text_center)
	ws.write('J8', 'Currency', f_text_center)

	ws.set_column(0, 0, 10)
	ws.set_column(1, 1, 15)
	ws.set_column(2, 9, 10)

	############ DATA ROWS ############
	row = 8

	for k in sqlSelectList(curs, sql2, (client_id, 'A', remit_id)):
	    i = tuple_to_clean_list(k)

	    ws.write(row, 0, i[0], f_text_normal)     # Account Number
	    ws.write(row, 1, i[1], f_text_normal)     # Debtor Name
	    ws.write(row, 2, i[2], f_text_normal)     # Payment Memo
	    ws.write(row, 3, i[3].strftime('%Y-%m-%d'), f_text_normal)     # Post Date
	    ws.write(row, 4, i[4], f_text_normal)     # Principal
	    ws.write(row, 5, i[5], f_text_normal)     # Commission
	    ws.write(row, 6, i[6], f_text_normal)     # Remit To
	    ws.write(row, 7, i[7], f_text_normal)     # Due From
	    ws.write(row, 8, i[8], f_text_normal)     # Total Due
	    ws.write(row, 9, currency, f_text_normal)     # Currency
	    if client_id in ['CISC19', 'CISC26', 'CISC38', 'CISC39', 'CISC40']:
	    	ws.write('K8', 'Type', f_text_center)
	    	if i[10] == 60:
	    		ws.write(row, 10, 'Credit', f_text_normal)
	    	else:
	    		ws.write(row, 10, 'Payment', f_text_normal)
	    	ws.write(row, 11, '.')
	    else:
	    	ws.write(row, 10, '.')

	    row += 1

	row += 1

	ws.merge_range(row, 0, row, 9, '')
	ws.write('A7', 'Paid To Client Detail', f_text_normal)
	ws.write('A8', 'Account Number', f_text_center)
	ws.write('B8', 'Debtor Number', f_text_center)
	ws.write('C8', 'Payment Memo', f_text_center)
	ws.write('D8', 'Post Date', f_text_center)
	ws.write('E8', 'Principal', f_text_center)
	ws.write('F8', 'Commission', f_text_center)
	ws.write('G8', 'Remit To', f_text_center)
	ws.write('H8', 'Due From', f_text_center)
	ws.write('I8', 'Total Due', f_text_center)
	ws.write('J8', 'Currency', f_text_center)

	for k in sqlSelectList(curs, sql2, (client_id, 'C', remit_id)):
	    i = tuple_to_clean_list(k)

	    ws.write(row, 0, i[0], f_text_normal)     # Account Number
	    ws.write(row, 1, i[1], f_text_normal)     # Debtor Name
	    ws.write(row, 2, i[2], f_text_normal)     # Payment Memo
	    ws.write(row, 3, i[3].strftime('%Y-%m-%d'), f_text_normal)     # Post Date
	    ws.write(row, 4, i[4], f_text_normal)     # Principal
	    ws.write(row, 5, i[5], f_text_normal)     # Commission
	    ws.write(row, 6, i[6], f_text_normal)     # Remit To
	    ws.write(row, 7, i[7], f_text_normal)     # Due From
	    ws.write(row, 8, i[8], f_text_normal)     # Total Due
	    ws.write(row, 9, currency, f_text_normal)     # Currency
	    if client_id in ['CISC19', 'CISC26', 'CISC38', 'CISC39', 'CISC40']:
	    	ws.write('K8', 'Type', f_text_center)
	    	if i[10] == 60:
	    		ws.write(row, 10, 'Credit', f_text_normal)
	    	else:
	    		ws.write(row, 10, 'Payment', f_text_normal)
	    	ws.write(row, 11, '.')
	    else:
	    	ws.write(row, 10, '.')

	    row += 1
	wb.close()
