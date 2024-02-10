import openpyxl
from openpyxl.styles import Border, Side, PatternFill
import string
import datetime
import json
import time
import copy
import warnings

# ปิดการแสดงข้อความแจ้งเตือนทั้งหมด
warnings.filterwarnings("ignore")


class SchedulePlanning:
    def __init__(
        self,
        prd_start_date,
    ):
        self.prd_start_date = prd_start_date
        self.current_value = []
        # find last row and get A4 to C at Last row
        self.remain = []
        self.ot_plan = False

    def reshape_data(self, data):
        reshaped_data = []
        for i in range(0, len(data), 5):
            batch = data[i : i + 5]
            reshaped_data.append(batch)
        return reshaped_data

    def paramChangeModel(self, prv_model, model):
        if model == prv_model:
            return 0
        else:
            return 1

    def getDayOfWeek(self, date_string):
        try:
            day_of_week = date_string.strftime("%A")  # หรือ %a หากต้องการรูปแบบที่สั้นกว่า
            return day_of_week
        except ValueError:
            return "Invalid date format"

    def getFirstPRD(self, data):
        unique_values = set(item[-1] for item in data)
        result = {}
        i = data[0][-1]

        for value in unique_values:
            for item in data:
                if item[-1] == value:
                    result[i] = item[4]
                    break
            i += 1
        return result

    def get_prd_line(self, data_list):
        unique_items = {}

        for item in data_list:
            last_element = item[-1]
            if last_element not in unique_items:
                unique_items[last_element] = last_element

        unique_result = unique_items.values()
        return unique_result

    def average(self, list_data):
        avg = 0
        if len(list_data) > 0:
            avg = sum(list_data) / len(list_data)
        return avg

    def is_datetime(self, value):
        return isinstance(value, datetime.datetime)

    def simulate_plan(self, original_data, man_power, prd_model, targetModel):
        data = copy.deepcopy(original_data)
        first_prd_lines = self.prd_start_date
        # เซ็ตค่าตัวเเปร
        ot = 0
        waiting_time = []
        sale_date_diff = []
        col = 0
        row = 0
        uph = 0
        i = 0
        round_ot = 0
        qty = 0
        actual_per_day = []
        sq_all = []
        prd_batch = []
        block_batch = []
        prd_date = []
        ot_batch = []
        batch_prd_date = []
        ot_interlock = []
        batch_index = None
        count = 0
        cmt_count = 0
        sequense = []
        msg = True
        lines = data[0][-1]
        prd_model_line = prd_model[lines]
        map_line = {1: "H1", 2: "H2", 3: "H3", 4: "H4"}
        current_date = first_prd_lines[lines]
        qty_batch = [item[3] for item in data]
        model_batch = [item[1] for item in data]
        set_batch = [item[0] for item in data]
        holiday = [
            datetime.datetime(2023, 12, 5, 0, 0),
            datetime.datetime(2023, 12, 30, 0, 0),
            datetime.datetime(2023, 12, 31, 0, 0),
            datetime.datetime(2024, 1, 1, 0, 0),
            datetime.datetime(2024, 1, 2, 0, 0),
            datetime.datetime(2024, 4, 12, 0, 0),
            datetime.datetime(2024, 4, 13, 0, 0),
            datetime.datetime(2024, 4, 15, 0, 0),
            datetime.datetime(2024, 4, 16, 0, 0),
            datetime.datetime(2024, 5, 1, 0, 0),
            datetime.datetime(2024, 5, 22, 0, 0),
            datetime.datetime(2024, 7, 20, 0, 0),
            datetime.datetime(2024, 7, 22, 0, 0),
            datetime.datetime(2024, 7, 29, 0, 0),
            datetime.datetime(2024, 8, 12, 0, 0),
            datetime.datetime(2024, 10, 14, 0, 0),
            datetime.datetime(2024, 12, 5, 0, 0),
            datetime.datetime(2024, 12, 30, 0, 0),
            datetime.datetime(2024, 12, 31, 0, 0),
        ]
        while row < len(data):
            qty = int(data[row][3])
            eta = data[row][4]
            model = data[row][1]
            due_date = (
                data[row][5]
                if self.is_datetime(data[row][5])
                else data[row][4] + datetime.timedelta(days=360)
            )
            waiting_time.append((current_date - eta).days)
            # print("Current Date:", current_date)
            # print("ETA:", eta)
            if model not in prd_model_line:
                # print("Condition Model")
                # print("model:", model)
                # print("Condition model:", prd_model_line)
                # time.sleep(3)
                return float("+inf"), float("+inf"), float("+inf"), [], {}
            while data[row][3] > 0:
                # print("Loop:", data[row][0], data[row][3])
                date_check = self.getDayOfWeek(current_date)
                uph = float(targetModel[data[row][2]]) * man_power
                if data[row][0] in ot_batch:
                    working_time = 10
                    if date_check != "Sunday":
                        ot += 1
                else:
                    working_time = 8
                upd = round((uph * working_time), 0)
                if current_date < (eta + datetime.timedelta(days=5)):
                    # print("BATCH:", data[row][0])
                    # print("Current Date:", current_date)
                    # print("ETA:", eta)
                    # print("ไม่สามารถจัดเเผนผลิตได้ เนื่องจาก ETA เกิน Current Date")
                    msg = False
                    # print("Result:", float("+inf"), float("+inf"), float("+inf"), [])
                    return float("+inf"), float("+inf"), float("+inf"), [], {}
                else:
                    # today is Sunday ?
                    if date_check == "Sunday" or current_date in holiday:
                        # shift 1 col
                        current_date = current_date + datetime.timedelta(days=1)
                    elif (
                        int(data[row][3]) == qty
                        and len(actual_per_day) > 0
                        and int(data[row][3]) <= upd
                        and actual_per_day[-1] == upd
                    ):
                        cmt_factor = self.paramChangeModel(
                            data[row - 1][1], data[row][1]
                        )
                        remainChangeBatch = (
                            round(
                                (int(data[row][3]) - actual_per_day[-1]) * cmt_factor,
                                0,
                            )
                            if int(data[row][3]) > actual_per_day[-1]
                            else round(
                                (actual_per_day[-1] - int(data[row][3])) * cmt_factor,
                                0,
                            )
                        )
                        if remainChangeBatch > int(data[row][3]):
                            remainChangeBatch = int(data[row][3])
                        actual_per_day.append(
                            round(abs(remainChangeBatch), 0)
                        )  # Actual
                        sq_all.append(f"{round(abs(remainChangeBatch), 0)}A0")
                        prd_batch.append(abs(remainChangeBatch))  # Actual
                        data[row][3] = (
                            int(data[row][3]) - remainChangeBatch
                            if remainChangeBatch > 0
                            else 0
                        )
                        sequense.append(f"{remainChangeBatch}:A0")
                        prd_date.append(current_date)
                        block_batch.append(data[row][0])
                        current_date = current_date + datetime.timedelta(days=1)
                    elif (row == 0 and int(data[row][3]) >= upd) or (
                        row != 0
                        and data[row - 1][-1] != data[row][-1]
                        and int(data[row][3]) >= upd
                    ):
                        data[row][3] -= upd
                        self.remain.append(int(data[row][3]))
                        actual_per_day.append(upd)  # Actual
                        sq_all.append(f"{upd}A1")
                        prd_batch.append(upd)  # Actual
                        count += 1
                        sequense.append(f"{upd}:A1")
                        prd_date.append(current_date)
                        current_date = current_date + datetime.timedelta(days=1)
                        block_batch.append(data[row][0])
                    # commond batch
                    elif (
                        len(self.remain) > 0
                        and self.remain[-1] < upd
                        and int(data[row][3]) > upd
                        and len(actual_per_day) > 0
                    ):
                        # get previous upd
                        prv_uph = float(targetModel[data[row - 1][2]]) * man_power
                        prv_upd = prv_uph * working_time
                        cmt_factor = self.paramChangeModel(
                            data[row - 1][1], data[row][1]
                        )
                        # count change model group by cabinet
                        if cmt_factor == 0.9:
                            cmt_count += 1
                        # check scarp
                        if actual_per_day[-1] < prv_upd:
                            # check joint change batch
                            if (
                                (
                                    (upd * cmt_factor) - self.remain[-1] <= upd
                                    or (upd * cmt_factor) - self.remain[-1] <= prv_upd
                                )
                                and date_check != "Monday"
                                and current_date - datetime.timedelta(days=1)
                                not in holiday
                            ):
                                col = 1
                            # shift 2 column for monday to sat
                            elif (
                                (
                                    (upd * cmt_factor) - self.remain[-1] <= upd
                                    or (upd * cmt_factor) - self.remain[-1] <= prv_upd
                                )
                                and date_check == "Monday"
                                and current_date - datetime.timedelta(days=2)
                                not in holiday
                            ):
                                col = 2
                            elif current_date - datetime.timedelta(days=1) in holiday:
                                current_date_copy = current_date - datetime.timedelta(
                                    days=1
                                )
                                col = 1
                                while True:
                                    if current_date_copy in holiday:
                                        current_date_copy -= datetime.timedelta(days=1)
                                        col += 1
                                    else:
                                        break
                            if data[row][0] in ot_batch:
                                working_time = 10
                                ot -= 0.5
                            else:
                                working_time = 8
                            upd = round((uph * working_time), 0)
                            prv_uph = float(targetModel[data[row - 1][2]]) * man_power
                            prv_upd = (
                                float(targetModel[data[row - 1][2]])
                                * man_power
                                * working_time
                            )
                            remain_hour_prv = abs(
                                round(prv_upd - (cmt_factor * prv_uph), 0) / prv_uph
                            ) - (round(actual_per_day[-1], 0) / prv_uph)
                            remainChangeBatch = round(uph * remain_hour_prv, 0)
                            # print("-------------------")
                            # print("prv_upd:", prv_upd)
                            # print("prv_upd * cmt_factor:", prv_upd * cmt_factor)
                            # print("upd * cmt_factor:", upd * cmt_factor)
                            # print("remain_hour_prv:", remain_hour_prv)
                            # print("actual_per_day[-1]:", actual_per_day[-1])
                            # print("remainChangeBatch:", remainChangeBatch)
                            actual_per_day.append(remainChangeBatch)
                            sq_all.append(f"{remainChangeBatch}C1")
                            prd_batch.append(remainChangeBatch)
                            if remainChangeBatch > 0:
                                count += round(remainChangeBatch / upd, 0)
                                sequense.append(f"{remainChangeBatch}C1")
                            # print("QTY:", data[row][3])
                            # print("remainChangeBatch:", remainChangeBatch)
                            data[row][3] -= remainChangeBatch
                            self.remain.append(data[row][3])
                            current_date = current_date - datetime.timedelta(days=col)
                            prd_date.append(current_date)
                            current_date = current_date + datetime.timedelta(days=col)
                            block_batch.append(data[row][0])
                        else:
                            data[row][3] -= upd
                            self.remain.append(int(data[row][3]))
                            actual_per_day.append(upd)  # Actual
                            sq_all.append(f"{upd}C2")
                            prd_batch.append(upd)  # Actual
                            count += 1
                            sequense.append(f"{upd}:C2")
                            prd_date.append(current_date)
                            current_date = current_date + datetime.timedelta(days=1)
                            block_batch.append(data[row][0])
                    elif int(data[row][3]) <= upd:
                        prv_uph = float(targetModel[data[row - 1][2]]) * man_power
                        prv_upd = prv_uph * working_time
                        if (
                            len(self.remain) > 0
                            and len(actual_per_day) > 0
                            and self.remain[-1] == 0
                            and actual_per_day[-1] < upd
                            and (int(data[row][3]) + actual_per_day[-1]) <= upd
                        ):
                            cmt_factor = self.paramChangeModel(
                                data[row - 1][1], data[row][1]
                            )
                            if cmt_factor == 1:
                                cmt_count += 1
                            if date_check != "Monday":
                                col = 1
                            else:
                                col = 2
                            if data[row][0] in ot_batch:
                                working_time = 10
                                ot -= 0.5
                            else:
                                working_time = 8
                            upd = round((uph * working_time * man_power), 0)
                            remainChangeBatch = round(
                                ((upd - (cmt_factor * uph)) - actual_per_day[-1]),
                                0,
                            )
                            actual_per_day.append(int(data[row][3]))  # Actual
                            sq_all.append(f"{int(data[row][3])}D1")
                            prd_batch.append(int(data[row][3]))  # Actual
                            sequense.append(f"{int(data[row][3])}D1")
                            data[row][3] = 0
                            count += int(data[row][3]) / upd
                            self.remain.append(int(data[row][3]))
                            current_date = current_date - datetime.timedelta(days=col)
                            prd_date.append(current_date)
                            current_date = current_date + datetime.timedelta(days=col)
                            block_batch.append(data[row][0])
                        elif (
                            len(self.remain) > 0
                            and len(actual_per_day) > 0
                            and self.remain[-1] == 0
                            and actual_per_day[-1] < upd
                        ):
                            if (
                                date_check != "Monday"
                                and current_date - datetime.timedelta(days=1)
                                not in holiday
                            ):
                                col = 1
                            elif current_date - datetime.timedelta(days=1) in holiday:
                                current_date_copy = current_date - datetime.timedelta(
                                    days=1
                                )
                                col = 1
                                while True:
                                    if current_date_copy in holiday:
                                        current_date_copy -= datetime.timedelta(days=1)
                                        col += 1
                                    else:
                                        break
                            else:
                                col = 2
                            actual_per_day.append(round(int(data[row][3]), 0))  # Actual
                            sq_all.append(f"{round(int(data[row][3]), 0)}D2")
                            prd_batch.append(int(data[row][3]))  # Actual
                            sequense.append(f"{data[row][3]}D2")
                            data[row][3] = 0
                            self.remain.append(0)
                            current_date = current_date - datetime.timedelta(days=col)
                            prd_date.append(current_date)
                            current_date = current_date + datetime.timedelta(
                                days=col + 1
                            )
                            block_batch.append(data[row][0])
                        else:
                            actual_per_day.append(round(int(data[row][3]), 0))  # Actual
                            sq_all.append(f"{round(int(data[row][3]), 0)}D3")
                            prd_batch.append(int(data[row][3]))  # Actual
                            sequense.append(f"{data[row][3]}D3")
                            data[row][3] = 0
                            self.remain.append(0)
                            prd_date.append(current_date)
                            current_date = current_date + datetime.timedelta(days=1)
                            block_batch.append(data[row][0])
                    elif int(data[row][3]) >= upd:
                        data[row][3] -= upd
                        actual_per_day.append(round(upd, 0))  # Actual
                        sq_all.append(f"{round(upd, 0)}E")
                        prd_batch.append(upd)  # Actual
                        self.remain.append(data[row][3])
                        count += 1
                        sequense.append(f"{upd}E")
                        prd_date.append(current_date)
                        current_date = current_date + datetime.timedelta(days=1)
                        block_batch.append(data[row][0])
                    # nomal put
                    else:
                        actual_per_day.append(round(int(data[row][3]), 0))  # Actual
                        sq_all.append(f"{round(int(data[row][3]), 0)}F")
                        sequense.append(f"{int(data[row][3])}F")
                        prd_batch.append(int(data[row][3]))  # Actual
                        data[row][3] = 0
                        prd_date.append(current_date)
                        current_date = current_date + datetime.timedelta(days=1)
                        count += 1
                        self.remain.append(0)
                        block_batch.append(data[row][0])

            if due_date - datetime.timedelta(days=3) < current_date:
                batch_index = set_batch.index(data[row][0])
                if ot_interlock.count(data[row][0]) > batch_index:
                    # print("Result:", float("+inf"), float("+inf"), float("+inf"), [])
                    return float("+inf"), float("+inf"), float("+inf"), [], {}
                elif len(ot_batch) == len(data):
                    # print("Result:", float("+inf"), float("+inf"), float("+inf"), [])
                    return float("+inf"), float("+inf"), float("+inf"), [], {}
                else:
                    ot_interlock.append(data[row][0])
                if len(ot_batch) == 0 and data[row][0] not in ot_batch:
                    ot_batch.append(data[row][0])
                    data[row][3] = qty_batch[row]
                    # print("OT Batch 1:", ot_batch)
                else:
                    row = batch_index - 1 if batch_index > 0 else 0
                    if row == 0:
                        ot_batch = [data[i][0] for i in range(len(data))]
                        row = 0
                        ot = 0
                        sale_date_diff = []
                        waiting_time = []
                        actual_per_day = []
                        prd_date = []
                        block_batch = []
                        current_date = first_prd_lines[lines]
                        for index in range(len(data)):
                            data[index][3] = qty_batch[index]
                    else:
                        # print("Row:", row)
                        # print("OT Batch:", ot_batch)
                        prd_batch = []
                        data[row][3] = qty_batch[row]
                        row = row - 1
                        (
                            ot_batch.append(data[row][0])
                            if data[row][0] not in ot_batch
                            else None
                        )
                        data[row][3] = qty_batch[row]
                    # print("OT Batch 2:", ot_batch)
                    # print(data[row][3], data[row][0])
                    # print("Block Batch:", block_batch)
                while len(block_batch) > 0 and block_batch[-1] in ot_batch:
                    ot -= 1
                    if actual_per_day:
                        actual_per_day.pop()
                    if block_batch:
                        block_batch.pop()
                    if prd_date:
                        prd_date.pop()
                    if sequense:
                        sequense.pop()
                    if len(prd_date) == 0:
                        current_date = first_prd_lines[lines]
                    else:
                        current_date = prd_date[-1] + datetime.timedelta(days=1)
                    # print("qty_batch:", data[row][0], qty_batch[row])
                    # print("actual_per_day:", actual_per_day)
                    # print("block_batch:", block_batch)
                    # print("prd_date:", prd_date)
                    # print("current_date:", current_date)
                    # print("due_date:", due_date)
                    # print("-------------OT------------")
                # print("OT Batch:", ot_batch)
            else:
                # print("OT:", ot)
                # print("Batch:", data[row][0])
                # print("Model:", data[row][1])
                # print("prd_batch:", prd_batch)
                # print("prd_date:", prd_date)
                # print("Sequence:", sequense)
                # print(f"Actual Per Day: {actual_per_day}")
                # print(f"Sequence: {sq_all}")
                # print("current_date:", current_date)
                # print("due_date:", due_date)
                # print("--------------------------")
                row += 1
                # self.ot_plan = False
                sale_date_diff.append((due_date - current_date).days)
            prd_batch = []
            sequense = []
        # print("Status :", msg)
        # print(f"Actual Per Day: {actual_per_day}")
        # print("prd_date:", prd_date)
        c = 0
        date_schedule = {}
        for qty, model in zip(qty_batch, model_batch):
            while c < len(actual_per_day):
                if qty != 0:
                    qty -= actual_per_day[c]
                    date_schedule[prd_date[c]] = model
                    c += 1
                else:
                    break
        print("Result:", ot, max(waiting_time), max(sale_date_diff))
        return (
            ot,
            max(waiting_time),
            max(sale_date_diff),
            ot_batch,
            date_schedule,
        )

    def return_data(self, data, line):
        i = 0
        for arry in data:
            arry.append(line[i])
            del arry[2]
            i += 1
        return data
