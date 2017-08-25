import traceback
import datetime
from os import sys

# Add the executing directory to the class path
sys.path.append("C:\\Users\\mscales\\Desktop\\Development\\MSSQLupdater\\")
from datetime import timedelta

from src.counter_function_library import (get_report_dates, open_client_report, run_client_report,
                                          make_head_node, create_output_report_xlsx)
from src.link_list import SingleLinkList
from src.CONSTANTS import SELF_PATH


def main(date_start=datetime.datetime.today(), date_end=None):
    if date_end is None:
        date_end = date_start
    running_date = date_start
    head_node = make_head_node()
    running_report = SingleLinkList(head_node)
    while date_end >= running_date:
        (client_report,
         report_page_rows,
         data_page_rows,
         report_page,
         data_page,
         report_date) = open_client_report(running_date)
        try:
            if client_report is None:
                raise FileNotFoundError
            else:
                client_report = run_client_report(client_report,
                                                  data_page_rows,
                                                  report_page_rows,
                                                  data_page,
                                                  report_page,
                                                  report_date)
        except FileNotFoundError:
            pass
        except Exception as err:
            now = datetime.datetime.now()
            error = traceback.format_exc()
            error_log = open(SELF_PATH + '\\MSSQLupdater\\error_logs\\%s_error_log.txt' %
                             report_date.strftime("%m%d%Y"), 'a')
            error_log.write('%r Exception Occurred: %s File Date: %r\n%s ' % (now.strftime("%m/%d/%Y %H:%M:%S"),
                                                                              err,
                                                                              report_date.strftime("%m/%d/%Y"),
                                                                              error))
            error_log.close()
            pass
        else:
            running_report.add_back(client_report)
        finally:
            running_date = running_date + timedelta(days=1)
    # running_report.push_to_sql()
    final_dict = running_report.make_accumulated_excel()
    create_output_report_xlsx(final_dict)


if __name__ == "__main__":
    import datetime
    import time
    from os import sys, path

    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
    fmt = '%m%d%Y'
    try:
        start_date = sys.argv[1]
        end_date = sys.argv[2]
        start_date = datetime.datetime.strptime(start_date, fmt)
        end_date = datetime.datetime.strptime(end_date, fmt)
    except IndexError:
        start_date = get_report_dates("What is the start date?")
        end_date = get_report_dates("What is the end date?")
    main(start_date, end_date)
else:
    from os import sys, path

    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
    main(datetime.datetime.now() - timedelta(days=1))
