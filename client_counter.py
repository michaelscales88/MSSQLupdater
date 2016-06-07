import traceback
import datetime
from datetime import timedelta
from link_list import SingleLinkList
from counter_function_library import get_report_dates, open_client_report, run_client_report
from CONSTANTS import SELF_PATH


def main():
    start_date = get_report_dates("What is the start date?")
    end_date = get_report_dates("What is the end date?")
    running_report = SingleLinkList()
    while end_date >= start_date:

        (client_report,
         report_page_rows,
         data_page_rows,
         report_page,
         data_page,
         report_date) = open_client_report(start_date)
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
            start_date = start_date + timedelta(days=1)
    running_report.push_to_sql()
    # running_report.print()


if __name__ == "__main__":
    main()
