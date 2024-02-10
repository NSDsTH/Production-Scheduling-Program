import datetime
from stock_INJ import Injection
import json


class StockManager:
    def __init__(self, date_schedule):
        self.model = []
        self.date_schedule = date_schedule
        self.inj_path = ""
        self.inj_sheet_name = ""
        self.lg = []
        self.hs = []

    def is_datetime(self, value):
        return isinstance(value, datetime.datetime)

    def check_sublists(self, list_test):
        found = 0
        supplier_logic = [
        ]
        if len(list_test) > 2:
            for sublist_logic in supplier_logic:
                if [list_test[0], list_test[1]] == sublist_logic:
                    found += 1
                if [list_test[1], list_test[2]] == sublist_logic:
                    found += 1
        else:
            for sublist_logic in supplier_logic:
                if list_test == sublist_logic:
                    found = 1
                    break
        return found

    def get_max_and_min_date(self):
        max_date = None
        min_date = None

        for schedule in self.date_schedule:
            current_max_date = max(schedule.keys(), default=None)
            current_min_date = min(schedule.keys(), default=None)

            if current_max_date and (not max_date or current_max_date > max_date):
                max_date = current_max_date

            if current_min_date and (not min_date or current_min_date < min_date):
                min_date = current_min_date
        return min_date, max_date

    def supplier_condition(self):
        count_list = 0
        start_date, end_date = self.get_max_and_min_date()
        if self.is_datetime(start_date) and self.is_datetime(end_date):
            while start_date <= end_date:
                h1 = self.date_schedule[0] if len(self.date_schedule[0]) > 0 else {}
                h2 = self.date_schedule[1] if len(self.date_schedule[1]) > 1 else {}
                h3 = self.date_schedule[2] if len(self.date_schedule) > 2 else {}
                prd = [h1.get(start_date), h2.get(start_date), h3.get(start_date)]
                count_list += self.check_sublists(prd)
                start_date += datetime.timedelta(days=1)
        return count_list
