# -*- coding: utf-8 -*-
import sys

sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python/')
from report_module import *
from format_only_module import *

seq = """
SELECT DBR_NO, DBR_CLI_REF_NO, DBR_NAME1, DATE_FORMAT(DAT_TRX_DATE_O, '%Y-%m-%d'), dat_seq_no,
(select GROUP_CONCAT(replace(DNT_NOTE, char(22), '''') ORDER BY DNT_SEQ_SUB SEPARATOR  '')
from CDS.DNT where dnt_dbr_no = DBR_NO and dnt_seq_no = dat_seq_no) as note
FROM CDS.DBR 
INNER JOIN CDS.DAT ON DAT_DBR_NO = DBR_NO
WHERE DBR_CLIENT in ('ECI097', 'ECI098', 'ECI099')
AND DBR_CLOSE_DATE_O IS NULL
AND DAT_ACTION_CODE in ('DPW','DGW','DCD','DCO','DES','DPO','DOF','DAC','DAR','DCA','DCH','DCM','DDB','DPF','DMN','DPI',
'DPR','DRT','DPD','DQU','DCT','DOA','DLF','DPA','DWL','DND','DTX')
AND DAT_TRX_DATE_O >= DATE_ADD(CURRENT_DATE, INTERVAL -7 DAY)
HAVING NOTE LIKE '{%'
order by DAT_TRX_DATE_O desc
"""

sql_select_invoice_due = """
SELECT SUM(IVT_AMOUNT - IVT_PAID) FROM CDS.IVT
WHERE IVT_DBR_NO = '%s'
AND IVT_IVT_NO = '%s'
"""

sql_invoice_date = """
select DATE_FORMAT(IVT_IVT_DATE_O ,'%%Y-%%m-%%d') from CDS.IVT
where ivt_dbr_no = '%s'
and ivt_ivt_no = '%s'
"""

sql_check_invoice_status = """
select ivt_ivt_no, sts_desc, DATE_FORMAT(ivt_status_date ,'%%Y-%%m-%%d') FROM CDS.IVT
INNER JOIN CDS.STS ON STS_CODE = IVT_STATUS
where ivt_dbr_no = '%s'
and ivt_ivt_no = '%s'
and ivt_paid - ivt_amount != 0
and ivt_status in ('DPW','DGW','DCD','DCO','DES','DPO','DOF','DAC','DAR','DCA','DCH','DCM','DDB','DPF','DMN','DPI',
'DPR','DRT','DPD','DQU','DCT','DOA','DLF','DPA','DWL','DND','DTX')
"""


def get_json_note(dragon_note):
    dragon_note = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', dragon_note)

    dragon_note = strip_non_ascii(dragon_note)

    json_note = json.loads((html_decode(dragon_note.decode('utf-8', 'ignore').encode('utf-8'))))

    return json_note


# Accepting the filename as an argument or using a variable to write it.
if "filename" not in jm:
    filename = 'ECI097_99_Dispute Report' + '-' + run_date.strftime('%Y-%m-%d') + '.xlsx'
    workbook = xlsxwriter.Workbook(filename)
else:
    workbook = xlsxwriter.Workbook(jm["filename"])

worksheet = workbook.add_worksheet()

f_text_left = workbook.add_format({'bold': True})
f_text_left.set_align('center')
f_text_left.set_font_color('#004c82')
f_text_left.set_font_size(16)

f_text_left_14 = workbook.add_format({'bold': True})
f_text_left_14.set_align('center')
f_text_left_14.set_font_color('#004c82')
f_text_left_14.set_font_size(12)

worksheet.set_column(0, 0, 16)
worksheet.set_column(1, 1, 19)
worksheet.set_column(2, 2, 30)
worksheet.set_column(3, 3, 16)
worksheet.set_column(4, 4, 25)
worksheet.set_column(5, 5, 20)
worksheet.set_column(6, 6, 25)
worksheet.set_column(7, 7, 25)
worksheet.set_column(8, 8, 40)
worksheet.set_column(9, 9, 40)
worksheet.set_column(10, 10, 30)

worksheet.merge_range(0, 0, 4, 1, add_logo(worksheet), ())
worksheet.merge_range(1, 3, 2, 6, "ECI097-ECI099 Dispute Report", f_text_left)
worksheet.merge_range(3, 3, 3, 6, run_date.strftime('%Y-%m-%d'), f_text_left_14)


f_text_left = workbook.add_format(text_left_top)
f_headers_format = workbook.add_format(headers_format)
f_text_wrap = workbook.add_format(text_left_top)
f_text_wrap.set_text_wrap()

worksheet.write('A7', 'Debtor Number ', f_headers_format)
worksheet.write('B7', 'Client Reference ', f_headers_format)
worksheet.write('C7', 'Name', f_headers_format)
worksheet.write('D7', 'Invoice Number', f_headers_format)
worksheet.write('E7', 'Invoice Status Description', f_headers_format)
worksheet.write('F7', 'Invoice Status Date', f_headers_format)
worksheet.write('G7', 'Contact Name', f_headers_format)
worksheet.write('H7', 'Contact Email ', f_headers_format)
worksheet.write('I7', 'Dispute Reason', f_headers_format)
worksheet.write('J7', 'Customer Request', f_headers_format)
worksheet.write('K7', 'Additional Notes', f_headers_format)

invoices_unique_list = []

row = 7
for l in sqlSelectList(curs, seq, ()):
    i = tuple_to_clean_list(l)

    note = get_json_note(i[5])

    list_invoices = note["widgetINVList"].split(',')

    seq_invoice_list = ""

    for elem in list_invoices:

        if elem not in invoices_unique_list:
            invoices_unique_list.append(elem)

            seq_invoice_list += elem + ","

    seq_invoice_list = seq_invoice_list[:-1]

    lista = list(seq_invoice_list.split(','))

    invoices_dispute_list = []

    for elem in lista:
        if len(elem) > 4:
            try:
                check_dispute = sqlSelectList(curs, sql_check_invoice_status, (i[0], elem))[0][0]
                invoices_dispute_list.append(check_dispute)
            except:
                pass

    for elemt in invoices_dispute_list:
        if len(elemt) > 4:
            total_due = sqlSelectList(curs, sql_select_invoice_due, (i[0], elemt))[0][0]
            try:
                check_dispute = sqlSelectList(curs, sql_check_invoice_status, (i[0], elemt))[0][0]
            except:
                check_dispute = elemt + ' - was in dispute'
        else:
            total_due = '0'

        try:
            total_due_amt = float(total_due)

            try:
                ivt_status_date = sqlSelectList(curs, sql_check_invoice_status, (i[0], elemt))[0][2]
            except:
                ivt_status_date = 'note placed on wrong account'

            try:
                sts_desc = sqlSelectList(curs, sql_check_invoice_status, (i[0], elemt))[0][1]
            except:
                sts_desc = ' '

            if total_due_amt != 0:

                worksheet.write(row, 0, i[0], f_text_left)
                worksheet.write(row, 1, i[1], f_text_left)
                worksheet.write(row, 2, i[2], f_text_left)
                worksheet.write(row, 3, check_dispute, f_text_left)
                worksheet.write(row, 4, sts_desc, f_text_left)
                worksheet.write(row, 5, ivt_status_date, f_text_left)
                worksheet.write(row, 6, note["contactName"], f_text_left)
                worksheet.write(row, 7, note["contactEmail"], f_text_left)

                try:
                    worksheet.write(row, 8, note["disputeReason"], f_text_wrap)
                except:
                    try:
                        worksheet.write(row, 8, note["debtDisputeReason"], f_text_wrap)
                    except:
                        worksheet.write(row, 8, '', f_text_wrap)

                try:
                    worksheet.write(row, 9, note["customerRequest"], f_text_wrap)
                except:
                    worksheet.write(row, 9, '', f_text_wrap)

                try:
                    worksheet.write(row, 10, note["additionalNotes"], f_text_wrap)
                except:
                    worksheet.write(row, 10, '', f_text_wrap)

                row += 1
        except:
            pass

workbook.close()