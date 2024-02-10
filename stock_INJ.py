import openpyxl
import datetime
import time


class Injection:
    def __init__(self, file_path, sheet_name):
        self.file_path = file_path
        self.sheet_name = sheet_name

    def get_data_from_sheet(self):
        if self.file_path != "":
            plan = openpyxl.load_workbook(self.file_path, data_only=True)
            sheet = plan[self.sheet_name]

            column_letters = ["A", "B", "C"]
            last_row = sheet.max_row
            i = 3
            data = []
            row = []
            while i < last_row:
                for col in column_letters:
                    cell_value = sheet[f"{col}{i}"].value
                    row.append(cell_value)
                data.append(row)
                row = []
                i = i + 1
            return data
        else:
            return []
