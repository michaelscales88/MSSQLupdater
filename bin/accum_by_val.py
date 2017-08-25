import traceback
import datetime
import pyexcel as pe
from collections import namedtuple, defaultdict
from os import sys

# Add the executing directory to the class path
sys.path.append("C:\\Users\\mscales\\Desktop\\Development\\MSSQLupdater\\")
from datetime import timedelta

from src.counter_function_library import (get_report_dates, open_client_report, run_client_report,
                                          make_head_node, create_output_report_xlsx, get_sec,
                                          convert_time_stamp)
from src.link_list import SingleLinkList
from src.CONSTANTS import SELF_PATH


class Counter:
    def __init__(self):
        self.morethan10 = []
        self.morethan5 = []
        self.lessthan5 = []


def main(date_start=datetime.datetime.today(), date_end=None):
    running_date = date_start
    # answered = defaultdict(list)
    # answered = namedtuple('answered', 'lessthan5 morethan5 morethan10')
    # answered(lessthan5=[], morethan5=[], morethan10=[])
    # # answered.__new__.__defaults__ = ([],) * len(answered._fields)
    # lost = namedtuple('lost', 'lessthan5 morethan5 morethan10')
    # lost(lessthan5=[], morethan5=[], morethan10=[])
    # # lost.__new__.__defaults__ = ([],) * len(lost._fields)
    # print(answered)
    # print(lost)
    answered = Counter()
    lost = Counter()
    while date_end >= running_date:
        month_date = running_date.strftime('%m%d%Y')
        # file = r'C:\Users\mscales\Desktop\Development\Attachment Archive\{0}\Call Details.xlsx'.format(month_date)
        file = r'C:\Users\mscales\Desktop\Development\Daily SLA Parser - Automated Version\Archive\{0}\{0}_Call Details.xlsx'.format(month_date)
        # print(file)
        try:
            sheet = pe.get_sheet(file_name=file)
            data = sheet.to_array()
            # print(data)
            for row in data:
                if row[6] == 7592:
                    call_duration = get_sec(row[12])
                    date = row[3]
                    if row[11] is True:
                        print("found true call")
                        if call_duration <= 30:
                            answered.morethan10.append(convert_time_stamp(call_duration) + '.' + date)
                        # elif call_duration >= 1800:
                        #     answered.morethan5.append(convert_time_stamp(call_duration) + '.' + date)
                        # else:
                        #     answered.lessthan5.append(convert_time_stamp(call_duration) + '.' + date)
                    # else:
                    #     if call_duration >= 600:
                    #         lost.morethan10.append(convert_time_stamp(call_duration) + '.' + date)
                    #     elif call_duration >= 300:
                    #         lost.morethan5.append(convert_time_stamp(call_duration) + '.' + date)
                    #     else:
                    #         lost.lessthan5.append(convert_time_stamp(call_duration) + '.' + date)
                else:
                    pass
        except Exception as e:
            print(e)
            pass
        running_date = running_date + timedelta(days=1)
    # print("Answered: less than 5 min: {}".format(len(answered.lessthan5)))
    # for duration in answered.lessthan5:
    #     print(duration)
    # print("Answered: more than 5 min: {}".format(len(answered.morethan5)))
    # for duration in answered.morethan5:
    #     print(duration)
    print("Answered: more than 60 min: {}".format(len(answered.morethan10)))
    for duration in answered.morethan10:
        print(duration.split('.'))
    # print("Lost: less than 5 min: {}".format(len(lost.lessthan5)))
    # for duration in lost.lessthan5:
    #     print(duration)
    # print("Lost: more than 5 min: {}".format(len(lost.morethan5)))
    # for duration in lost.morethan5:
    #     print(duration)
    # print("Lost: more than 10 min: {}".format(len(lost.morethan10)))
    # for duration in lost.morethan10:
    #     print(duration)

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
