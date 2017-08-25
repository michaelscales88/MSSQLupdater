class ClientContainer:
    def __init__(self, client_name):
        self.name = client_name
        self.report_date = 0
        self.total_calls_presented = 0
        self.total_calls_answered = 0
        self.total_calls_lost = 0
        self.voice_mails_received = 0
        self.average_incoming_duration = 0
        self.average_wait_answered = 0
        self.average_wait_lost = 0
        self.ticker_less_than_15_sec = 0
        self.ticker_less_than_30_sec = 0
        self.ticker_less_than_45_sec = 0
        self.ticker_less_than_60_sec = 0
        self.ticker_greater_than_60_sec = 0
        self.unique_id = 0

    # Setters
    def set_stats_for_client(self,
                             report_date,
                             total_calls_presented,
                             total_calls_lost,
                             voice_mails_received,
                             average_incoming_duration,
                             average_wait_answered,
                             average_wait_lost,
                             ticker_less_than_15_sec,
                             ticker_less_than_30_sec,
                             ticker_less_than_45_sec,
                             ticker_less_than_60_sec,
                             ticker_greater_than_60_sec):
        self.report_date = report_date
        self.total_calls_presented = total_calls_presented
        self.total_calls_answered = (total_calls_presented - total_calls_lost - voice_mails_received)
        self.total_calls_lost = total_calls_lost
        self.voice_mails_received = voice_mails_received
        self.average_incoming_duration = average_incoming_duration
        self.average_wait_answered = average_wait_answered
        self.average_wait_lost = average_wait_lost
        self.ticker_less_than_15_sec = ticker_less_than_15_sec
        self.ticker_less_than_30_sec = ticker_less_than_30_sec
        self.ticker_less_than_45_sec = ticker_less_than_45_sec
        self.ticker_less_than_60_sec = ticker_less_than_60_sec
        self.ticker_greater_than_60_sec = ticker_greater_than_60_sec
        self.unique_id = self.report_date.strftime('%m%d%Y') + self.name

    def set_stats_for_excel(self,
                            total_calls_presented,
                            total_calls_answered,
                            total_calls_lost,
                            voice_mails_received,
                            average_incoming_duration,
                            average_wait_answered,
                            average_wait_lost,
                            ticker_less_than_15_sec,
                            ticker_less_than_30_sec,
                            ticker_less_than_45_sec,
                            ticker_less_than_60_sec,
                            ticker_greater_than_60_sec):
        self.total_calls_presented += total_calls_presented
        self.total_calls_answered += total_calls_answered
        self.total_calls_lost += total_calls_lost
        self.voice_mails_received += voice_mails_received
        self.average_incoming_duration += average_incoming_duration
        self.average_wait_answered += average_wait_answered
        self.average_wait_lost += average_wait_lost
        self.ticker_less_than_15_sec += ticker_less_than_15_sec
        self.ticker_less_than_30_sec += ticker_less_than_30_sec
        self.ticker_less_than_45_sec += ticker_less_than_45_sec
        self.ticker_less_than_60_sec += ticker_less_than_60_sec
        self.ticker_greater_than_60_sec += ticker_greater_than_60_sec

    # Getters
    def get_sql_string(self):
        return (self.name,
                self.total_calls_presented,
                self.total_calls_answered,
                self.total_calls_lost,
                self.voice_mails_received,
                self.average_incoming_duration,
                self.average_wait_answered,
                self.average_wait_lost,
                self.ticker_less_than_15_sec,
                self.ticker_less_than_30_sec,
                self.ticker_less_than_45_sec,
                self.ticker_less_than_60_sec,
                self.ticker_greater_than_60_sec,
                self.report_date,
                self.unique_id)

    def get_name(self):
        return self.name

    def get_stats_for_excel(self):
        return (self.total_calls_presented,
                self.total_calls_answered,
                self.total_calls_lost,
                self.voice_mails_received,
                self.average_incoming_duration,
                self.average_wait_answered,
                self.average_wait_lost,
                self.ticker_less_than_15_sec,
                self.ticker_less_than_30_sec,
                self.ticker_less_than_45_sec,
                self.ticker_less_than_60_sec,
                self.ticker_greater_than_60_sec)

    def get_voice_mails(self):
        return self.voice_mails_received

    def get_total_calls(self):
        return self.total_calls_presented

    def get_true_calls_lost(self):
        return self.total_calls_lost

    def get_average_duration(self):
        return self.average_incoming_duration

    def get_average_wait_answered(self):
        return self.average_wait_answered

    def get_average_wait_lost(self):
        return self.average_wait_lost

    # def get_longest_answered(self):
    #     return self.longest_answered_time

    def get_ticker(self):
        return (self.ticker_less_than_15_sec,
                self.ticker_less_than_30_sec,
                self.ticker_less_than_45_sec,
                self.ticker_less_than_60_sec,
                self.ticker_greater_than_60_sec)

    # Special getters for custom output
    def get_calls_answered_true(self):
        return self.total_calls_answered

    def get_calls_answered_percentage(self):
        answered_percentage = 0
        try:
            answered_percentage = (self.get_calls_answered_true() / self.get_total_calls()) * 100
            answered_percentage = round(answered_percentage, 2)
        except ZeroDivisionError:
            answered_percentage = 0
        finally:
            return "{0}%".format(answered_percentage)

    def get_calls_lost_percentage(self):
        lost_percentage = 0
        try:
            lost_percentage = (self.get_true_calls_lost() / self.get_total_calls()) * 100
            lost_percentage = round(lost_percentage, 2)
        except ZeroDivisionError:
            lost_percentage = 0
        finally:
            return "{0}%".format(lost_percentage)

    def get_pca_percentage(self):
        pca_percentage = 0
        try:
            pca_percentage = ((self.ticker_less_than_15_sec + self.ticker_less_than_30_sec) /
                              self.get_calls_answered_true()) * 100
            pca_percentage = round(pca_percentage, 2)
        except ZeroDivisionError:
            pca_percentage = 0
        finally:
            return "{0}%".format(pca_percentage)

    # Utilities
    def __str__(self):
        print('report_date {}'.format(self.report_date))
        print('client_name %s' % self.name)
        print('total_calls_presented %d' % self.total_calls_presented)
        print('total_calls_answered %d' % self.total_calls_answered)
        print('total_calls_lost %d' % self.total_calls_lost)
        print('voice_mails_received %d' % self.voice_mails_received)
        print('average_incoming_duration %d' % self.average_incoming_duration)
        print('average_wait_answered %d' % self.average_wait_answered)
        print('average_wait_lost %d' % self.average_wait_lost)
        print('ticker_less_than_15_sec: %d' % self.ticker_less_than_15_sec)
        print('ticker_less_than_30_sec: %d' % self.ticker_less_than_30_sec)
        print('ticker_less_than_45_sec %d' % self.ticker_less_than_45_sec)
        print('ticker_less_than_60_sec %d' % self.ticker_less_than_60_sec)
        print('ticker_greater_than_60_sec %d' % self.ticker_greater_than_60_sec)

    def write_out_file(self, text_file):
        text_file.write('report_date {}\n'.format(self.report_date))
        text_file.write('client_name %s\n' % self.name)
        text_file.write('total_calls_presented %d\n' % self.total_calls_presented)
        text_file.write('total_calls_answered %d\n' % self.total_calls_answered)
        text_file.write('total_calls_lost %d\n' % self.total_calls_lost)
        text_file.write('voice_mails_received %d\n' % self.voice_mails_received)
        text_file.write('average_incoming_duration %d\n' % self.average_incoming_duration)
        text_file.write('average_wait_answered %d\n' % self.average_wait_answered)
        text_file.write('average_wait_lost %d\n' % self.average_wait_lost)
        text_file.write('ticker_less_than_15_sec: %d\n' % self.ticker_less_than_15_sec)
        text_file.write('ticker_less_than_30_sec: %d\n' % self.ticker_less_than_30_sec)
        text_file.write('ticker_less_than_45_sec %d\n' % self.ticker_less_than_45_sec)
        text_file.write('ticker_less_than_60_sec %d\n' % self.ticker_less_than_60_sec)
        text_file.write('ticker_greater_than_60_sec %d\n' % self.ticker_greater_than_60_sec)
