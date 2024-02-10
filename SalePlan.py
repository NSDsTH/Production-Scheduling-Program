import openpyxl
import datetime
import json


class SalePlan:
    def __init__(self, filename, sheetname):
        self.filename = filename
        self.sheet_name = sheetname

    def calculate_date_difference(self, date1, date2):
        date_format = "%Y/%m/%d"
        try:
            # แปลงวันที่ให้เป็นออบเจ็กต์ datetime
            date1 = datetime.datetime.strptime(date1, date_format).date()
            date2 = datetime.datetime.strptime(date2, date_format).date()

            # คำนวณผลต่างของวันที่
            diff = (date2 - date1).days

            return diff
        except ValueError:
            return "Invalid date format"

    def check_data_type(self, data):
        checked_data = []
        for item in data:
            # ดึงข้อมูลสูตรทางคณิตศาสตร์
            formula = item[-1]
            if isinstance(formula, str):
                # ดำเนินการคำนวณค่าสูตรทางคณิตศาสตร์และแปลงเป็นตัวเลข
                result = eval(formula.strip("="))
                # เปลี่ยนค่าสูตรทางคณิตศาสตร์ในลิสต์เป็นผลลัพธ์ที่คำนวณแล้ว
                item[-1] = int(result)

            elif isinstance(formula, int):
                pass
            checked_data.append(item)
        return checked_data

    def sort_data_by_date(self, data):
        # Cut the array at the last index where it is None
        block = []
        for sublist in data:
            if all(item is not None for item in sublist):
                block.append(sublist)

        return block

    def get_data_from_sheet(self):
        plan = openpyxl.load_workbook(self.filename, data_only=True)
        sheet = plan[self.sheet_name]

        # Define the column letters
        # B = BATCH, C = MODEL, A = Loading date
        column_letters = ["B", "C", "N"]
        # Determine the last row with information
        last_row = sheet.max_row
        block_batch = []
        i = 2
        data = []
        row = []
        while i < last_row:
            for col in column_letters:
                cell_value = sheet[f"{col}{i}"]
                row.append(cell_value.value)
            if None not in row and row[0] not in block_batch:
                data.append(row)
                block_batch.append(row[0])
            row = []
            i += 1
        return self.sort_data_by_date(data)

    def process_data(self):
        data_sheet = self.get_data_from_sheet()
        data = self.check_data_type(data_sheet)
        return data
