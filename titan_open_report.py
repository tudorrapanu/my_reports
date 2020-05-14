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
select
dbr_client,
dbr_no,
dbr_cli_ref_no,
dbr_name1,
dbr_assign_date_o,
dbr_assign_amt,
dbr_assign_amt - dbr_recvd_tot as balance
from cds.dbr
where 
dbr_client in ('TTMI82', 'TTMI83')
and dbr_class not in (2, 3)
"""

notes_headers = (["Client ID", "Debtor Number","Client Reference", "Customer Name", "Assigned Date", "Assigned Amount", "Balance"])

write_report(sql_select_notes,'client_portal_notes',notes_headers,'Titan Open Accounts Report')