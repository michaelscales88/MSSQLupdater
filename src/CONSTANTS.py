import os
from os import path

SELF_PATH = os.path.dirname(path.dirname(path.abspath(__file__)))

OUTPUT_PATH = '\\Daily SLA Parser - Automated Version\\Output\\'
M_DRIVE_PATH = 'M:\\Help Desk\\Daily SLA Report\\'
CONNECTION_STRING = "Driver={SQL Server};Server=localhost;Database=mw_calling;Trusted_Connection=yes"
SQL_COMMAND = ('''INSERT INTO sla_summary(client_name,
               calls_presented,
               calls_answered,
               calls_lost,
               voice_mails,
               average_incoming_duration,
               average_wait_answered,
               average_wait_lost,
               calls_less_15,
               calls_less_30,
               calls_less_45,
               calls_less_60,
               calls_greater_60,
               report_date,
               unique_id)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''')
