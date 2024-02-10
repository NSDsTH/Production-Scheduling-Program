import openpyxl
import datetime
import time


class OrderPlan:
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name

    def sort_data_by_date(self, data):
        # Cut the array at the last index where it is None
        block = []
        for sublist in data:
            if all(item is not None for item in sublist):
                if isinstance(sublist[-1], str) and "\n" in sublist[-1]:
                    sublist[-1] = sublist[-1].replace("\n", "")
                    date_format = "%Y/%m/%d"
                    sublist[-1] = datetime.datetime.strptime(sublist[-1], date_format)
                block.append(sublist)

        return block

    def get_data_from_sheet(self):
        if self.file_path != "":
            plan = openpyxl.load_workbook(self.file_path, data_only=True)
            sheet = plan[self.sheet_name]

            column_letters = ["C", "D", "H", "I", "X"]
            last_row = sheet.max_row
            i = 3
            data = []
            row = []
            while i < last_row:
                for col in column_letters:
                    cell_value = sheet[f"{col}{i}"].value
                    row.append(cell_value)
                if "大货无膜片-空运" not in row and None not in row:
                    data.append(row)
                row = []
                i = i + 1
            return self.sort_data_by_date(data)
        else:
            return []
