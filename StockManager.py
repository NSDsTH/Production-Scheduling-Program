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

    # def stock_supplier(self):
    #     stock_all = {}
    #     stock_injection = {model_name: float("inf") for model_name in self.model}
    #     stock_store = {model_name: float("inf") for model_name in self.model}

    #     # order = Injection(self.inj_path, self.inj_sheet_name)
    #     # stock_inj = order.get_data_from_sheet()
    #     stock_inj = [
    #         [39000245, "1983829  Fan snail shell Y1K,LG-Y1K", 900],
    #         [39000246, "1983831  Fan snail shell Y1K,LG-Y1K", 753],
    #         [39000247, "1983832  Indoor fan volute Y1K,LG-Y1K", 948],
    #         [39000248, "1983833  Indoor fan volute Y1K,LG-Y1K", 2584],
    #         [39000249, "1983836  Back Board Y1K,LG-Y1K", 100],
    #         [39000250, "1983844  Dummy coupling Y1K/Y1N/Y1L", 2331],
    #         [39000251, "1983845  Dummy coupling Y1K/Y1N/Y1L", 2025],
    #         [39000252, "1983851  Level vane Y1K,LG-Y1K", 8000],
    #         [39000253, "1983852  Vertical louver Y1K,LG-Y1K", 12000],
    #         [39000255, "1983860  Filter Frame Y1K,LG-Y1K", 3240],
    #         [39000256, "1984945Indoor fan volute Y1LCW,HW,LG-Y1L", 3304],
    #         [39000260, "1989641  Shell pattern Y1K", 4234],
    #         [39000261, "1989706  Base holder Y1K,LG-Y1K", 2291],
    #         [39000262, "2133201  Air outlet Y1N", 792],
    #         [39000263, "2133202  Level vane Y1N", 530],
    #         [39000264, "2133208  Water tray Y1N,LG-Y1N", 800],
    #         [39000265, "2133209  Fan snail shell Y1N,LG-Y1N", 876],
    #         [39000266, "2133210  Fan snail shell Y1N,LG-Y1N", 2268],
    #         [39000267, "2133211  Indoor fan volute Y1N,LG-Y1N", 699],
    #         [39000268, "2133212  Indoor fan volute Y1N,LG-Y1N", 1800],
    #         [39000274, "2137621  Air outlet  Y1P", 2475],
    #         [39000276, "2137813  Base Y1P", 756],
    #         [39000279, "2137851  Fan snail shell Y1P", 477],
    #         [39000281, "2137853  Water tray Y1P", 1440],
    #         [39000283, "2137855  Indoor fan volute Y1P", 789],
    #         [39000288, "2142325  Base holder part Y1N,LG-Y1N", 1116],
    #         [39000289, "2142553  Panel pattern Y1N", 48],
    #         [39000290, "2147584  Panel pattern Y1P", 396],
    #         [39000293, "2152468  Base Y1L CW,HW,TW,LG-Y1L", 1192],
    #         [39000294, "2152491Fan snail Y1L CW,HW,TW,LG-Y1L", 1074],
    #         [39000295, "2152492Fan snail Y1L CW,HW,TW,LG-Y1L", 1349],
    #         [39000300, "2190603  Back board Y1N", 419],
    #         [39000301, "2190604  Filter Frame Y1N", 2120],
    #         [39000357, "1984946  Back Board Y1L CW,HW,TW,AP700", 507],
    #         [39000361, "2210024  Receiving window Y1L CW,HW,TW", 1865],
    #         [39000362, "2210026  Level vane Y1L ( CW,HW,TW )", 4951],
    #         [39000365, "1984950  Filter Frame Y1L ( TW )", 2232],
    #         [39000370, "1892493  Fan shell  Y1G,Y1L CW,HW,LG-Y1L", 503],
    #         [39000371, "1984947  Filter Frame Y1L CW,HW,TW,AP700", 314],
    #         [39000398, "1892495 SUPPORTER BASE Y1G", 563],
    #         [39000400, "1892492 Fan shell  Y1G", 44],
    #         [39000401, "1892494 Fan shell  Y1G", 408],
    #         [39000403, "2227622 Back board Y1G 8K,HP  ( White )", 963],
    #         [39000405, "1892499 Display case panel 8K,HP Y1G ( W", 636],
    #         [39000406, "1892490 level vane assy Y1G 8K,HP", 1799],
    #         [39000407, "1892502 Display case window Y1G 8K,HP", 1120],
    #         [39000410, "2227685 Filter Frame Y1G", 1260],
    #         [39000420, "2287674 Panel 6K LG-Y1N", 1106],
    #         [39000421, "2161445 Back Board 6K,7K LG-Y1N", 129],
    #         [39000423, "2162511 Filter Frame 6K,7K LG-Y1N", 1603],
    #         [39000425, "2287606 Leave vane 6K,7K LG-Y1N", 2384],
    #         [39000426, "2287607 Leave vane 6K,7K LG-Y1N", 368],
    #         [39000427, "2290914 Leave vane 8K LG-Y1N", 1228],
    #     ]

    #     with open("Stock.json", "r") as json_file:
    #         stores_on_hand = json.load(json_file)

    #     for item in stock_inj:
    #         part_name = item[1]
    #         quantity = item[2]
    #         for model_name in self.model:
    #             if model_name in part_name and not any(
    #                 model in part_name for model in self.lg
    #             ):
    #                 stock_injection[model_name] = min(
    #                     stock_injection[model_name], quantity
    #                 )
    #             elif model_name in part_name and any(
    #                 model in part_name for model in self.lg
    #             ):
    #                 stock_injection[model_name] = min(
    #                     stock_injection[model_name], quantity
    #                 )

    #     for item in stores_on_hand:
    #         part_name = item[0]
    #         quantity = item[1]
    #         for model_name in self.model:
    #             if model_name in part_name and not any(
    #                 model in part_name for model in self.lg
    #             ):
    #                 stock_store[model_name] = min(stock_store[model_name], quantity)
    #             elif model_name in part_name and any(
    #                 model in part_name for model in self.lg
    #             ):
    #                 stock_store[model_name] = min(stock_store[model_name], quantity)

    #     for model_inj, model_store in zip(stock_injection, stock_store):
    #         if (
    #             stock_injection[model_inj] == float("inf")
    #             and stock_store[model_store] == float("inf")
    #             and model_inj == model_store
    #         ):
    #             stock_all[model_store] = 0
    #         else:
    #             stock_all[model_store] = (
    #                 stock_injection[model_inj] + stock_store[model_store]
    #             )

    #     return stock_all

    # def add_values_in_dict(self, input_dict):
    #     cap_upd = {
    #         "Y1G": 1100,
    #         "Y1K": 924,
    #         "Y1L": 880,
    #         "Y1N": 880,
    #         "Y1P": 770,
    #         "LG-Y1K": 924,
    #         "LG-Y1L": 880,
    #         "LG-Y1N": 880,
    #         "C1A": 0,
    #         "C1B": 0,
    #         "C1F": 0,
    #         "C1U": 0,
    #         "C1V": 0,
    #         "C1W": 0,
    #         "C1Y": 0,
    #         "C2X": 0,
    #         "C3D": 0,
    #     }
    #     for key in input_dict:
    #         input_dict[key] += cap_upd[key]
    #     return input_dict

    def is_datetime(self, value):
        return isinstance(value, datetime.datetime)

    def check_sublists(self, list_test):
        found = 0
        supplier_logic = [
            ["Y1G", "Y1G"],
            ["C1W", "C1W"],
            ["Y1L", "Y1L"],
            ["Y1K", "Y1K"],
            ["C1F", "C1V"],
            ["Y1L", "Y1G"],
            ["02", "02"],
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


# date_schedule = [
#     {},
#     {
#         datetime.datetime(2023, 12, 2, 0, 0): "Y1K",
#         datetime.datetime(2023, 12, 4, 0, 0): "Y1K",
#         datetime.datetime(2023, 12, 6, 0, 0): "Y1K",
#         datetime.datetime(2023, 12, 7, 0, 0): "Y1K",
#         datetime.datetime(2023, 12, 8, 0, 0): "Y1K",
#         datetime.datetime(2023, 12, 9, 0, 0): "Y1K",
#         datetime.datetime(2023, 12, 11, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 12, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 13, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 14, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 15, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 16, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 18, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 19, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 20, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 21, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 22, 0, 0): "Y1N",
#         datetime.datetime(2023, 12, 23, 0, 0): "Y1N",
#     },
# ]

# stock_manager = StockManager(date_schedule)
# stock_manager = stock_manager.supplier_condition()
# print("Stock Manager:", stock_manager)
