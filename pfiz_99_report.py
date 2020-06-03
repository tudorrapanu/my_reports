# -*- coding: utf-8 -*-

from __future__ import unicode_literals
import sys
sys.path.append('/home/build/scoop/server/python/')
sys.path.append('/python')
from report_module import *
from format_only_module import *
import datetime
from dateutil.relativedelta import relativedelta

sql_select_notes = """
select dbr_client,
dbr_name1,
dbr_cli_ref_no,
(select isnull(sum(ivt_amount-ivt_paid), 0) 
        from cds.ivt 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=0) as bucket0, 
(select isnull(sum(ivt_amount-ivt_paid), 0) 
        from cds.ivt 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=30) as bucket30,
(select isnull(sum(ivt_amount-ivt_paid), 0)  
        from cds.ivt 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=60) as bucket60, 
(select isnull(sum(ivt_amount-ivt_paid), 0)  
        from cds.ivt 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=90) as bucket90, 
(select isnull(sum(ivt_amount-ivt_paid), 0) 
        from cds.ivt 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=120) as bucket120,
(select isnull(sum(ivt_amount-ivt_paid), 0) 
        from cds.ivt 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=180) as bucket180,
(select isnull(sum(ivt_amount-ivt_paid), 0)  
        from cds.ivt 
        where ivt_dbr_no = dbr_no 
        and inv_past_due_bucket=181) as bucket181,  
dbr_status, 
dbr_last_worked_o,
(select max(dat_Seq_no) 
        from cds.dat 
        where dat_dbr_no = dbr_no 
        and dat_note = 'Y') as max_seq,
(select list(dnt_note, '' order by dnt_seq_sub) 
        from cds.dnt 
        where dnt_dbr_no = dbr_no 
        and dnt_seq_no = max_seq) as last_note
from cds.dbr
where dbr_client = 'PFIZ99'
group by
dbr_status, 
dbr_last_worked_o,
dbr_client,
dbr_no,
dbr_name1,
dbr_cli_ref_no;
"""

notes_headers = (["Client ID", "Client Name","Client Reference", "<0 Days", "0-30 Days", "31-60 Days", "61-90 Days", "90-120 Days", "121-180 Days", "181+ Days", "Account Status", "Last Worked Date", "Max Sequence", "Last Note"])

write_report(sql_select_notes,'pfiz_99_report',notes_headers,'PFIZ99 Accounts Report')