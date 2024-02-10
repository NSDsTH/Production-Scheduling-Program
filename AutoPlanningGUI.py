import sys
import os
import datetime
import time
import subprocess
import re
import json
import threading
import openpyxl
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QMainWindow
from PyQt5.QtCore import QTime, QDate, QThread, pyqtSignal
import numpy as np
from SalePlan import SalePlan
from OrderPlan import OrderPlan
from DataPlan import DataPlan
from PlanningProcess import PlanningProcess
from ClearData import ClearDataPlanning
from MultiProcesing import MultiProcessingScheduler


def ProductionPlanning(self):
    self.status.setText("Processing")
    with open("parameter.json", "r") as json_file:
        data = json.load(json_file)
    data["status"] = "Processing"
    with open("parameter.json", "w") as json_file:
        json.dump(data, json_file)

    file_path = self.data_json["file_path"]
    sheet_name = self.data_json["prd_sheet_name"]
    order_path = self.data_json["order_path"]
    order_sheet_name = self.data_json["order_sheet_name"]
    sale_path = self.data_json["sale_path"]
    sale_sheet_name = self.data_json["sale_sheet_name"]

    def get_first_plan(order_data, sale_data):
        data_planning = []
        for order in order_data:
            swit = 0
            for sale in sale_data:
                if order[0] == sale[-4]:
                    order_with_sale = order + [sale[0]]
                    data_planning.append(order_with_sale)
                    break
            if swit == 0:
                order_with_dash = order + ["-"]
                data_planning.append(order_with_dash)
        return data_planning

    def merge_lists(list1, list2):
        unique_data = set(map(tuple, list1))
        unique_data.update(map(tuple, list2))
        merged_list = [list(item) for item in unique_data]
        return merged_list

    def get_con_plan(current_plan, order_data, sale_data):
        data_planning = []
        union_list = merge_lists(current_plan, order_data)
        for order in union_list:
            swit = 0
            for sale in sale_data:
                if order[0] == sale[0]:
                    order_with_sale = order + [sale[-1]]
                    data_planning.append(order_with_sale)
                    break
            if swit == 0:
                order_with_dash = order + ["-"]
                data_planning.append(order_with_dash)
        return data_planning

    def datetime_handler(x):
        if isinstance(x, datetime.datetime):
            return x.timestamp()
        raise TypeError("Unknown type")

    def sort_key(item):
        return item[-1]

    processor = SalePlan(sale_path, sale_sheet_name)
    sale_data = processor.process_data()

    order = OrderPlan(order_path, order_sheet_name)
    order_data = order.get_data_from_sheet()

    data_get_plan = DataPlan(file_path, sheet_name)
    current_plan, current_row = data_get_plan.getRemainBatch()
    data_planning = (
        get_first_plan(order_data, sale_data)
        if self.data_json["status_plan"] == "NEW SEASON"
        else get_con_plan(current_plan, order_data, sale_data)
    )

    map_line = {"HISENSE 1": 1, "HISENSE 2": 2, "HISENSE 3": 3, "HISENSE 4": 4}
    prd_start_date = {
        1: datetime.datetime.strptime(self.data_json["prd_start_h1"], "%Y-%m-%d"),
        2: datetime.datetime.strptime(self.data_json["prd_start_h2"], "%Y-%m-%d"),
        3: datetime.datetime.strptime(self.data_json["prd_start_h3"], "%Y-%m-%d"),
        4: datetime.datetime.strptime(self.data_json["prd_start_h4"], "%Y-%m-%d"),
    }
    status_plan = self.data_json["status_plan"]
    ot_plan = self.data_json["ot_plan"]
    (
        current_column,
        current_value,
        list_start_date,
    ) = data_get_plan.find_last_value_column_in_row(current_row)
    prd_line = [map_line[line] for line in self.data_json["prd_line"]]
    population_line = int(self.data_json["population_line"])
    population_size = int(self.data_json["population_size"])
    generations = int(self.data_json["generations"])
    mutation_rate = float(self.data_json["mutation_rate"])
    man_power_h1 = int(self.data_json["man_power_h1"])
    man_power_h2 = int(self.data_json["man_power_h2"])
    man_power_h3 = int(self.data_json["man_power_h3"])
    man_power_h4 = int(self.data_json["man_power_h4"])
    checkbox_h1 = self.data_json["checkbox_h1"]
    checkbox_h2 = self.data_json["checkbox_h2"]
    checkbox_h3 = self.data_json["checkbox_h3"]
    checkbox_h4 = self.data_json["checkbox_h4"]
    num_processes = int(self.data_json["core_processing"])
    man_power = {1: man_power_h1, 2: man_power_h2, 3: man_power_h3, 4: man_power_h4}
    prd_model = {1: checkbox_h1, 2: checkbox_h2, 3: checkbox_h3, 4: checkbox_h4}
    targetModel = self.data_json["UPH"]
    # รวมวันที่ เข้าด้วยกัน
    combined_list_date = {**prd_start_date, **list_start_date}
    # Create and run the ProductionScheduler
    print("---------- Data Planning ----------")
    for i in data_planning:
        print(i)
    scheduler = MultiProcessingScheduler(
        data_planning,
        prd_line,
        population_line,
        population_size,
        generations,
        mutation_rate,
        combined_list_date,
        num_processes,
        man_power=man_power,
        prd_model=prd_model,
        targetModel=targetModel,
    )
    production_schedule, objective_prd_schedule = scheduler.run()
    production_schedule = sorted(production_schedule, key=sort_key)
    print("---------- Production Schedule ----------")
    for batch in production_schedule:
        print(batch)
    Process_planning = PlanningProcess(
        file_path,
        sheet_name,
        status_plan,
        prd_start_date,
        current_column,
        current_value,
        current_row,
        ot_plan,
        targetModel,
    )
    # best_ot_score, best_setup_time, best_inventory_day, best_sale_date_diff
    print("---Production Schedule Complete---")
    ot = objective_prd_schedule[0]
    change_model = objective_prd_schedule[1]
    inventory_day = objective_prd_schedule[2]
    sale_date_diff = objective_prd_schedule[3]
    msg = True
    if (
        ot != float("+inf")
        and change_model != float("+inf")
        and inventory_day != float("+inf")
    ):
        status = ""
        data_put_target = Process_planning.putTarget(production_schedule, man_power)
        # Write Parameter
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["status"] = "Complete"
        data["change_model"] = str(change_model)
        data["total_ot_day"] = str(round(ot, 0))
        data["follow_sale_plan"] = status
        data["supplier_follow"] = status
        data["date_diff"] = str(inventory_day)
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        print("Complete")
        self.status.setText("Complete")
        self.change_model.setText(str(change_model))
        self.date_diff.setText(str(inventory_day))
        # Write Result Planning
        plan_result = json.dumps(data_put_target, default=datetime_handler)
        with open("ResultPlanning.json", "w") as json_file:
            json_file.write(plan_result)
    else:
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["status"] = "Failed"
        data["change_model"] = str(0)
        data["total_ot_day"] = str(0)
        data["follow_sale_plan"] = "Failed"
        data["supplier_follow"] = "Failed"
        data["date_diff"] = str(0)
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        print("Failed")
        self.status.setText("Failed")
        return False
            

class ProductionPlanningThread(QThread):
    result_ready = pyqtSignal(str)
    error_occurred = pyqtSignal(str)  # เพิ่ม signal สำหรับข้อความผิดพลาด

    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

    def run(self):
        try:
            run_prd = ProductionPlanning(self.main_window)
            if run_prd == False:
                self.error_occurred.emit("Unable to schedule production")
            else:
                self.result_ready.emit("Finish Processing")
        except Exception as e:
            self.error_occurred.emit(str(e))  # ส่ง signal หากเกิดข้อผิดพลาด

class Ui_MainWindow(QMainWindow, object):
    def __init__(self):
        super().__init__()
        self.data_json = {}
        self.load_data()
        self.production_thread = ProductionPlanningThread(self)
        self.production_thread.result_ready.connect(self.display_result)
        self.production_thread.error_occurred.connect(self.showErrorAlert)

    def load_data(self):
        try:
            with open("parameter.json", "r") as json_file:
                self.data_json = json.load(json_file)
                self.checkbox_h1 = self.data_json["checkbox_h1"]
                self.checkbox_h2 = self.data_json["checkbox_h2"]
                self.checkbox_h3 = self.data_json["checkbox_h3"]
                self.checkbox_h4 = self.data_json["checkbox_h4"]
                self.prd_line = self.data_json["prd_line"]
                self.ot_plan = self.data_json["ot_plan"]
        except FileNotFoundError:
            self.data_json = {}

    def fetch_data_periodically(self, interval_seconds):
        while True:
            self.load_data()
            time.sleep(interval_seconds)

    def start_heavy_process(self):
        self.pushButton_4.setEnabled(False)
        self.production_thread.start()

    def display_result(self, result):
        self.pushButton_4.setEnabled(True)
        # self.status.setText("Complete")

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1176, 888)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Maximum
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_10 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_10.setObjectName("gridLayout_10")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setStyleSheet("background-color: rgb(0, 0, 0);")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.frame)
        self.horizontalLayout.setContentsMargins(-1, -1, -1, 4)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.frame)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setStyleSheet("color: rgb(255, 255, 255);")
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.gridLayout_10.addWidget(self.frame, 0, 0, 1, 1)
        self.widget_13 = QtWidgets.QWidget(self.centralwidget)
        self.widget_13.setMinimumSize(QtCore.QSize(0, 25))
        self.widget_13.setObjectName("widget_13")
        self.Home = QtWidgets.QPushButton(self.widget_13)
        self.Home.setGeometry(QtCore.QRect(2, 0, 91, 28))
        self.Home.setObjectName("Home")
        self.config = QtWidgets.QPushButton(self.widget_13)
        self.config.setGeometry(QtCore.QRect(100, 0, 91, 28))
        self.config.setObjectName("config")
        self.uph = QtWidgets.QPushButton(self.widget_13)
        self.uph.setGeometry(QtCore.QRect(200, 0, 93, 28))
        self.uph.setObjectName("uph")
        self.gridLayout_10.addWidget(self.widget_13, 1, 0, 1, 1)
        self.stackedWidget = QtWidgets.QStackedWidget(self.centralwidget)
        self.stackedWidget.setMinimumSize(QtCore.QSize(0, 760))
        self.stackedWidget.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.stackedWidget.setLineWidth(0)
        self.stackedWidget.setObjectName("stackedWidget")
        self.page_1 = QtWidgets.QWidget()
        self.page_1.setObjectName("page_1")
        self.gridLayout_6 = QtWidgets.QGridLayout(self.page_1)
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.widget = QtWidgets.QWidget(self.page_1)
        self.widget.setMinimumSize(QtCore.QSize(500, 0))
        self.widget.setMaximumSize(QtCore.QSize(5, 16777215))
        self.widget.setStyleSheet("background-color: rgb(240, 240, 240);")
        self.widget.setObjectName("widget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.widget)
        self.gridLayout_2.setContentsMargins(-1, 0, -1, -1)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.label_5 = QtWidgets.QLabel(self.widget)
        self.label_5.setMaximumSize(QtCore.QSize(16777215, 100))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.gridLayout_2.addWidget(self.label_5, 0, 0, 1, 2)
        self.label_6 = QtWidgets.QLabel(self.widget)
        self.label_6.setMinimumSize(QtCore.QSize(0, 0))
        self.label_6.setObjectName("label_6")
        self.gridLayout_2.addWidget(self.label_6, 3, 0, 1, 1)
        self.file_path = QtWidgets.QLineEdit(self.widget)
        font = QtGui.QFont()
        font.setPointSize(6)
        self.file_path.setFont(font)
        self.file_path.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.file_path.setText("")
        self.file_path.setObjectName("file_path")
        self.gridLayout_2.addWidget(self.file_path, 2, 0, 1, 3)
        self.label_2 = QtWidgets.QLabel(self.widget)
        self.label_2.setMaximumSize(QtCore.QSize(16777215, 50))
        self.label_2.setObjectName("label_2")
        self.gridLayout_2.addWidget(self.label_2, 1, 0, 1, 2)
        spacerItem = QtWidgets.QSpacerItem(
            20, 50, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed
        )
        self.gridLayout_2.addItem(spacerItem, 17, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.widget)
        self.label_3.setMaximumSize(QtCore.QSize(16777215, 50))
        self.label_3.setObjectName("label_3")
        self.gridLayout_2.addWidget(self.label_3, 7, 0, 1, 2)
        self.label_8 = QtWidgets.QLabel(self.widget)
        self.label_8.setMinimumSize(QtCore.QSize(0, 0))
        self.label_8.setObjectName("label_8")
        self.gridLayout_2.addWidget(self.label_8, 15, 0, 1, 1)
        self.prd_sheet_name = QtWidgets.QLineEdit(self.widget)
        self.prd_sheet_name.setMinimumSize(QtCore.QSize(10, 0))
        self.prd_sheet_name.setMaximumSize(QtCore.QSize(300, 16777215))
        self.prd_sheet_name.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.prd_sheet_name.setObjectName("prd_sheet_name")
        self.gridLayout_2.addWidget(self.prd_sheet_name, 3, 1, 1, 1)
        self.pushButton_3 = QtWidgets.QPushButton(self.widget)
        self.pushButton_3.setObjectName("pushButton_3")
        self.gridLayout_2.addWidget(self.pushButton_3, 13, 3, 1, 2)
        self.pushButton_2 = QtWidgets.QPushButton(self.widget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout_2.addWidget(self.pushButton_2, 8, 3, 1, 2)
        self.sale_path = QtWidgets.QLineEdit(self.widget)
        font = QtGui.QFont()
        font.setPointSize(6)
        self.sale_path.setFont(font)
        self.sale_path.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.sale_path.setObjectName("sale_path")
        self.gridLayout_2.addWidget(self.sale_path, 13, 0, 1, 3)
        self.order_path = QtWidgets.QLineEdit(self.widget)
        font = QtGui.QFont()
        font.setPointSize(6)
        self.order_path.setFont(font)
        self.order_path.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.order_path.setObjectName("order_path")
        self.gridLayout_2.addWidget(self.order_path, 8, 0, 1, 3)
        self.pushButton = QtWidgets.QPushButton(self.widget)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout_2.addWidget(self.pushButton, 2, 3, 1, 2)
        self.label_4 = QtWidgets.QLabel(self.widget)
        self.label_4.setMaximumSize(QtCore.QSize(16777215, 50))
        self.label_4.setObjectName("label_4")
        self.gridLayout_2.addWidget(self.label_4, 12, 0, 1, 2)
        self.order_sheet_name = QtWidgets.QLineEdit(self.widget)
        self.order_sheet_name.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.order_sheet_name.setObjectName("order_sheet_name")
        self.gridLayout_2.addWidget(self.order_sheet_name, 10, 1, 1, 1)
        self.label_7 = QtWidgets.QLabel(self.widget)
        self.label_7.setMinimumSize(QtCore.QSize(0, 0))
        self.label_7.setObjectName("label_7")
        self.gridLayout_2.addWidget(self.label_7, 10, 0, 1, 1)
        self.sale_sheet_name = QtWidgets.QLineEdit(self.widget)
        self.sale_sheet_name.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.sale_sheet_name.setObjectName("sale_sheet_name")
        self.gridLayout_2.addWidget(self.sale_sheet_name, 15, 1, 1, 1)
        self.gridLayout_8 = QtWidgets.QGridLayout()
        self.gridLayout_8.setSpacing(0)
        self.gridLayout_8.setObjectName("gridLayout_8")
        self.pushButton_4 = QtWidgets.QPushButton(self.widget)
        self.pushButton_4.setMinimumSize(QtCore.QSize(100, 0))
        self.pushButton_4.setMaximumSize(QtCore.QSize(300, 16777215))
        self.pushButton_4.setObjectName("pushButton_4")
        self.gridLayout_8.addWidget(self.pushButton_4, 0, 0, 1, 1)
        self.pushButton_5 = QtWidgets.QPushButton(self.widget)
        self.pushButton_5.setMinimumSize(QtCore.QSize(100, 0))
        self.pushButton_5.setMaximumSize(QtCore.QSize(300, 16777215))
        self.pushButton_5.setIconSize(QtCore.QSize(30, 20))
        self.pushButton_5.setObjectName("pushButton_5")
        self.gridLayout_8.addWidget(self.pushButton_5, 0, 1, 1, 1)
        self.gridLayout_2.addLayout(self.gridLayout_8, 19, 0, 1, 5)
        self.gridLayout_6.addWidget(self.widget, 0, 1, 1, 1)
        self.widget_2 = QtWidgets.QWidget(self.page_1)
        self.widget_2.setMinimumSize(QtCore.QSize(600, 0))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.widget_2.setFont(font)
        self.widget_2.setStyleSheet("background-color: rgb(240, 240, 240);")
        self.widget_2.setObjectName("widget_2")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.widget_2)
        self.gridLayout_3.setContentsMargins(-1, 11, -1, -1)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.label_17 = QtWidgets.QLabel(self.widget_2)
        self.label_17.setObjectName("label_17")
        self.gridLayout_3.addWidget(self.label_17, 5, 0, 1, 1)
        self.label_10 = QtWidgets.QLabel(self.widget_2)
        self.label_10.setObjectName("label_10")
        self.gridLayout_3.addWidget(self.label_10, 7, 2, 1, 1)
        self.label_11 = QtWidgets.QLabel(self.widget_2)
        self.label_11.setMinimumSize(QtCore.QSize(0, 100))
        self.label_11.setMaximumSize(QtCore.QSize(16777215, 100))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setKerning(False)
        self.label_11.setFont(font)
        self.label_11.setObjectName("label_11")
        self.gridLayout_3.addWidget(self.label_11, 0, 0, 1, 1)
        self.label_15 = QtWidgets.QLabel(self.widget_2)
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label_15.setFont(font)
        self.label_15.setObjectName("label_15")
        self.gridLayout_3.addWidget(self.label_15, 3, 0, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.widget_2)
        self.label_9.setObjectName("label_9")
        self.gridLayout_3.addWidget(self.label_9, 7, 0, 1, 1)
        self.label_13 = QtWidgets.QLabel(self.widget_2)
        self.label_13.setObjectName("label_13")
        self.gridLayout_3.addWidget(self.label_13, 8, 0, 1, 1)
        self.label_19 = QtWidgets.QLabel(self.widget_2)
        self.label_19.setObjectName("label_19")
        self.gridLayout_3.addWidget(self.label_19, 6, 0, 1, 1)
        self.label_12 = QtWidgets.QLabel(self.widget_2)
        self.label_12.setObjectName("label_12")
        self.gridLayout_3.addWidget(self.label_12, 8, 2, 1, 1)
        self.label_18 = QtWidgets.QLabel(self.widget_2)
        self.label_18.setObjectName("label_18")
        self.gridLayout_3.addWidget(self.label_18, 6, 2, 1, 1)
        self.status = QtWidgets.QLabel(self.widget_2)
        self.status.setMaximumSize(QtCore.QSize(200, 100))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.status.setFont(font)
        self.status.setStyleSheet(
            "background-color: rgb(255, 255, 255);\n" "border-color: rgb(0, 0, 0);"
        )
        self.status.setFrameShadow(QtWidgets.QFrame.Plain)
        self.status.setLineWidth(-1)
        self.status.setTextFormat(QtCore.Qt.AutoText)
        self.status.setAlignment(QtCore.Qt.AlignCenter)
        self.status.setWordWrap(False)
        self.status.setObjectName("status")
        self.gridLayout_3.addWidget(self.status, 0, 1, 1, 1)
        self.follow_sale_plan = QtWidgets.QLabel(self.widget_2)
        self.follow_sale_plan.setMaximumSize(QtCore.QSize(16777215, 30))
        self.follow_sale_plan.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.follow_sale_plan.setText("")
        self.follow_sale_plan.setAlignment(QtCore.Qt.AlignCenter)
        self.follow_sale_plan.setObjectName("follow_sale_plan")
        self.gridLayout_3.addWidget(self.follow_sale_plan, 3, 1, 1, 1)
        self.supplier_follow = QtWidgets.QLabel(self.widget_2)
        self.supplier_follow.setMaximumSize(QtCore.QSize(16777215, 30))
        self.supplier_follow.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.supplier_follow.setText("")
        self.supplier_follow.setAlignment(QtCore.Qt.AlignCenter)
        self.supplier_follow.setObjectName("supplier_follow")
        self.gridLayout_3.addWidget(self.supplier_follow, 5, 1, 1, 1)
        self.total_ot_day = QtWidgets.QLabel(self.widget_2)
        self.total_ot_day.setMaximumSize(QtCore.QSize(16777215, 30))
        self.total_ot_day.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.total_ot_day.setText("")
        self.total_ot_day.setAlignment(QtCore.Qt.AlignCenter)
        self.total_ot_day.setObjectName("total_ot_day")
        self.gridLayout_3.addWidget(self.total_ot_day, 6, 1, 1, 1)
        self.change_model = QtWidgets.QLabel(self.widget_2)
        self.change_model.setMaximumSize(QtCore.QSize(16777215, 30))
        self.change_model.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.change_model.setText("")
        self.change_model.setAlignment(QtCore.Qt.AlignCenter)
        self.change_model.setObjectName("change_model")
        self.gridLayout_3.addWidget(self.change_model, 7, 1, 1, 1)
        self.date_diff = QtWidgets.QLabel(self.widget_2)
        self.date_diff.setMaximumSize(QtCore.QSize(16777215, 30))
        self.date_diff.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.date_diff.setText("")
        self.date_diff.setAlignment(QtCore.Qt.AlignCenter)
        self.date_diff.setObjectName("date_diff")
        self.gridLayout_3.addWidget(self.date_diff, 8, 1, 1, 1)
        self.gridLayout_6.addWidget(self.widget_2, 0, 2, 1, 1)
        self.stackedWidget.addWidget(self.page_1)
        self.page_3 = QtWidgets.QWidget()
        self.page_3.setObjectName("page_3")
        self.gridLayout_7 = QtWidgets.QGridLayout(self.page_3)
        self.gridLayout_7.setObjectName("gridLayout_7")
        self.widget_3 = QtWidgets.QWidget(self.page_3)
        self.widget_3.setMinimumSize(QtCore.QSize(500, 0))
        self.widget_3.setStyleSheet("background-color: rgb(240, 240, 240);")
        self.widget_3.setObjectName("widget_3")
        self.gridLayout_5 = QtWidgets.QGridLayout(self.widget_3)
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.C1U_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C1U_h3.setObjectName("C1U_h3")
        self.gridLayout_5.addWidget(self.C1U_h3, 12, 1, 1, 1)
        self.C3D_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C3D_h4.setObjectName("C3D_h4")
        self.gridLayout_5.addWidget(self.C3D_h4, 16, 9, 1, 1)
        self.Y1N_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1N_h3.setObjectName("Y1N_h3")
        self.gridLayout_5.addWidget(self.Y1N_h3, 11, 5, 1, 1)
        self.C1W_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C1W_h3.setObjectName("C1W_h3")
        self.gridLayout_5.addWidget(self.C1W_h3, 12, 5, 1, 1)
        self.C1V_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C1V_h3.setObjectName("C1V_h3")
        self.gridLayout_5.addWidget(self.C1V_h3, 12, 3, 1, 1)
        self.Y1P_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1P_h3.setObjectName("Y1P_h3")
        self.gridLayout_5.addWidget(self.Y1P_h3, 11, 6, 1, 1)
        self.Friday = QtWidgets.QCheckBox(self.widget_3)
        self.Friday.setObjectName("Friday")
        self.gridLayout_5.addWidget(self.Friday, 19, 6, 1, 1)
        self.LGY1K_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1K_h3.setObjectName("LGY1K_h3")
        self.gridLayout_5.addWidget(self.LGY1K_h3, 13, 1, 1, 1)
        self.W1V_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.W1V_h1.setObjectName("W1V_h1")
        self.gridLayout_5.addWidget(self.W1V_h1, 5, 6, 1, 1)
        self.Saturday = QtWidgets.QCheckBox(self.widget_3)
        self.Saturday.setObjectName("Saturday")
        self.gridLayout_5.addWidget(self.Saturday, 19, 8, 1, 1)
        self.C2X_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C2X_h2.setObjectName("C2X_h2")
        self.gridLayout_5.addWidget(self.C2X_h2, 8, 8, 1, 1)
        self.C1B_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C1B_h2.setObjectName("C1B_h2")
        self.gridLayout_5.addWidget(self.C1B_h2, 7, 9, 1, 1)
        self.C3D_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C3D_h2.setObjectName("C3D_h2")
        self.gridLayout_5.addWidget(self.C3D_h2, 8, 9, 1, 1)
        self.C3D_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C3D_h3.setObjectName("C3D_h3")
        self.gridLayout_5.addWidget(self.C3D_h3, 12, 9, 1, 1)
        self.C2X_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C2X_h3.setObjectName("C2X_h3")
        self.gridLayout_5.addWidget(self.C2X_h3, 12, 8, 1, 1)
        self.C1B_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C1B_h3.setObjectName("C1B_h3")
        self.gridLayout_5.addWidget(self.C1B_h3, 11, 9, 1, 1)
        self.C1B_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C1B_h4.setObjectName("C1B_h4")
        self.gridLayout_5.addWidget(self.C1B_h4, 15, 9, 1, 1)
        self.C1A_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C1A_h3.setObjectName("C1A_h3")
        self.gridLayout_5.addWidget(self.C1A_h3, 11, 8, 1, 1)
        self.C1A_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C1A_h2.setObjectName("C1A_h2")
        self.gridLayout_5.addWidget(self.C1A_h2, 7, 8, 1, 1)
        self.Y1L_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1L_h3.setObjectName("Y1L_h3")
        self.gridLayout_5.addWidget(self.Y1L_h3, 11, 3, 1, 1)
        self.C1Y_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C1Y_h3.setObjectName("C1Y_h3")
        self.gridLayout_5.addWidget(self.C1Y_h3, 12, 6, 1, 1)
        self.Tuesday = QtWidgets.QCheckBox(self.widget_3)
        self.Tuesday.setObjectName("Tuesday")
        self.gridLayout_5.addWidget(self.Tuesday, 19, 1, 1, 1)
        self.W1R_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.W1R_h3.setObjectName("W1R_h3")
        self.gridLayout_5.addWidget(self.W1R_h3, 13, 5, 1, 1)
        self.W1M_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.W1M_h1.setObjectName("W1M_h1")
        self.gridLayout_5.addWidget(self.W1M_h1, 5, 3, 1, 1)
        self.Wednesday = QtWidgets.QCheckBox(self.widget_3)
        self.Wednesday.setObjectName("Wednesday")
        self.gridLayout_5.addWidget(self.Wednesday, 19, 3, 1, 1)
        self.W1R_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.W1R_h1.setObjectName("W1R_h1")
        self.gridLayout_5.addWidget(self.W1R_h1, 5, 5, 1, 1)
        self.label_26 = QtWidgets.QLabel(self.widget_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_26.setFont(font)
        self.label_26.setObjectName("label_26")
        self.gridLayout_5.addWidget(self.label_26, 2, 0, 1, 6)
        self.Thursday = QtWidgets.QCheckBox(self.widget_3)
        self.Thursday.setObjectName("Thursday")
        self.gridLayout_5.addWidget(self.Thursday, 19, 5, 1, 1)
        self.Y1G_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1G_h2.setObjectName("Y1G_h2")
        self.gridLayout_5.addWidget(self.Y1G_h2, 7, 0, 1, 1)
        self.Y1K_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1K_h1.setObjectName("Y1K_h1")
        self.gridLayout_5.addWidget(self.Y1K_h1, 3, 1, 1, 1)
        self.label_25 = QtWidgets.QLabel(self.widget_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_25.setFont(font)
        self.label_25.setObjectName("label_25")
        self.gridLayout_5.addWidget(self.label_25, 0, 0, 1, 6)
        self.Y1G_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1G_h1.setObjectName("Y1G_h1")
        self.gridLayout_5.addWidget(self.Y1G_h1, 3, 0, 1, 1)
        self.Y1G_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1G_h3.setObjectName("Y1G_h3")
        self.gridLayout_5.addWidget(self.Y1G_h3, 11, 0, 1, 1)
        self.C1F_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C1F_h1.setObjectName("C1F_h1")
        self.gridLayout_5.addWidget(self.C1F_h1, 4, 0, 1, 1)
        self.Y1K_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1K_h3.setObjectName("Y1K_h3")
        self.gridLayout_5.addWidget(self.Y1K_h3, 11, 1, 1, 1)
        self.Y1K_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1K_h2.setObjectName("Y1K_h2")
        self.gridLayout_5.addWidget(self.Y1K_h2, 7, 1, 1, 1)
        self.Y1L_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1L_h2.setObjectName("Y1L_h2")
        self.gridLayout_5.addWidget(self.Y1L_h2, 7, 3, 1, 1)
        self.Y1N_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1N_h2.setObjectName("Y1N_h2")
        self.gridLayout_5.addWidget(self.Y1N_h2, 7, 5, 1, 1)
        self.label_27 = QtWidgets.QLabel(self.widget_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_27.setFont(font)
        self.label_27.setObjectName("label_27")
        self.gridLayout_5.addWidget(self.label_27, 6, 0, 1, 6)
        self.Y1G_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1G_h4.setObjectName("Y1G_h4")
        self.gridLayout_5.addWidget(self.Y1G_h4, 15, 0, 1, 1)
        self.h4 = QtWidgets.QCheckBox(self.widget_3)
        self.h4.setObjectName("h4")
        self.gridLayout_5.addWidget(self.h4, 1, 9, 1, 2)
        self.C1V_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C1V_h2.setObjectName("C1V_h2")
        self.gridLayout_5.addWidget(self.C1V_h2, 8, 3, 1, 1)
        self.C1W_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C1W_h2.setObjectName("C1W_h2")
        self.gridLayout_5.addWidget(self.C1W_h2, 8, 5, 1, 1)
        self.Y1P_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1P_h2.setObjectName("Y1P_h2")
        self.gridLayout_5.addWidget(self.Y1P_h2, 7, 6, 1, 1)
        self.label_21 = QtWidgets.QLabel(self.widget_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_21.setFont(font)
        self.label_21.setObjectName("label_21")
        self.gridLayout_5.addWidget(self.label_21, 18, 0, 1, 4)
        self.label_28 = QtWidgets.QLabel(self.widget_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_28.setFont(font)
        self.label_28.setObjectName("label_28")
        self.gridLayout_5.addWidget(self.label_28, 10, 0, 1, 6)
        self.label_24 = QtWidgets.QLabel(self.widget_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_24.setFont(font)
        self.label_24.setObjectName("label_24")
        self.gridLayout_5.addWidget(self.label_24, 21, 0, 1, 4)
        self.C1F_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C1F_h4.setObjectName("C1F_h4")
        self.gridLayout_5.addWidget(self.C1F_h4, 16, 0, 1, 1)
        self.C1W_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C1W_h1.setObjectName("C1W_h1")
        self.gridLayout_5.addWidget(self.C1W_h1, 4, 5, 1, 1)
        self.Monday = QtWidgets.QCheckBox(self.widget_3)
        self.Monday.setObjectName("Monday")
        self.gridLayout_5.addWidget(self.Monday, 19, 0, 1, 1)
        self.new_season = QtWidgets.QCheckBox(self.widget_3)
        self.new_season.setObjectName("new_season")
        self.gridLayout_5.addWidget(self.new_season, 23, 0, 1, 4)
        self.W1V_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.W1V_h2.setObjectName("W1V_h2")
        self.gridLayout_5.addWidget(self.W1V_h2, 9, 6, 1, 1)
        self.W1M_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.W1M_h2.setObjectName("W1M_h2")
        self.gridLayout_5.addWidget(self.W1M_h2, 9, 3, 1, 1)
        self.Y1N_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1N_h4.setObjectName("Y1N_h4")
        self.gridLayout_5.addWidget(self.Y1N_h4, 15, 5, 1, 1)
        self.W1V_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.W1V_h3.setObjectName("W1V_h3")
        self.gridLayout_5.addWidget(self.W1V_h3, 13, 6, 1, 1)
        self.C1Y_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C1Y_h2.setObjectName("C1Y_h2")
        self.gridLayout_5.addWidget(self.C1Y_h2, 8, 6, 1, 1)
        self.C1F_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.C1F_h3.setObjectName("C1F_h3")
        self.gridLayout_5.addWidget(self.C1F_h3, 12, 0, 1, 1)
        self.C1F_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C1F_h2.setObjectName("C1F_h2")
        self.gridLayout_5.addWidget(self.C1F_h2, 8, 0, 1, 1)
        self.C1U_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C1U_h4.setObjectName("C1U_h4")
        self.gridLayout_5.addWidget(self.C1U_h4, 16, 1, 1, 1)
        self.Y1L_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1L_h4.setObjectName("Y1L_h4")
        self.gridLayout_5.addWidget(self.Y1L_h4, 15, 3, 1, 1)
        self.C1U_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.C1U_h2.setObjectName("C1U_h2")
        self.gridLayout_5.addWidget(self.C1U_h2, 8, 1, 1, 1)
        self.LGY1K_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1K_h2.setObjectName("LGY1K_h2")
        self.gridLayout_5.addWidget(self.LGY1K_h2, 9, 1, 1, 1)
        self.C1V_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C1V_h1.setObjectName("C1V_h1")
        self.gridLayout_5.addWidget(self.C1V_h1, 4, 3, 1, 1)
        self.C1V_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C1V_h4.setObjectName("C1V_h4")
        self.gridLayout_5.addWidget(self.C1V_h4, 16, 3, 1, 1)
        self.continue_plan = QtWidgets.QCheckBox(self.widget_3)
        self.continue_plan.setObjectName("continue_plan")
        self.gridLayout_5.addWidget(self.continue_plan, 22, 0, 1, 6)
        self.label_29 = QtWidgets.QLabel(self.widget_3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_29.setFont(font)
        self.label_29.setObjectName("label_29")
        self.gridLayout_5.addWidget(self.label_29, 14, 0, 1, 6)
        self.C3D_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C3D_h1.setObjectName("C3D_h1")
        self.gridLayout_5.addWidget(self.C3D_h1, 4, 9, 1, 1)
        self.C1B_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C1B_h1.setObjectName("C1B_h1")
        self.gridLayout_5.addWidget(self.C1B_h1, 3, 9, 1, 1)
        self.C2X_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C2X_h4.setObjectName("C2X_h4")
        self.gridLayout_5.addWidget(self.C2X_h4, 16, 8, 1, 1)
        self.Y1P_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1P_h1.setObjectName("Y1P_h1")
        self.gridLayout_5.addWidget(self.Y1P_h1, 3, 6, 1, 1)
        self.LGY1N_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1N_h2.setObjectName("LGY1N_h2")
        self.gridLayout_5.addWidget(self.LGY1N_h2, 9, 0, 1, 1)
        self.LGY1K_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1K_h4.setObjectName("LGY1K_h4")
        self.gridLayout_5.addWidget(self.LGY1K_h4, 17, 1, 1, 1)
        self.C1W_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C1W_h4.setObjectName("C1W_h4")
        self.gridLayout_5.addWidget(self.C1W_h4, 16, 5, 1, 1)
        self.C1U_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C1U_h1.setObjectName("C1U_h1")
        self.gridLayout_5.addWidget(self.C1U_h1, 4, 1, 1, 1)
        self.W1R_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.W1R_h2.setObjectName("W1R_h2")
        self.gridLayout_5.addWidget(self.W1R_h2, 9, 5, 1, 1)
        self.LGY1N_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1N_h4.setObjectName("LGY1N_h4")
        self.gridLayout_5.addWidget(self.LGY1N_h4, 17, 0, 1, 1)
        self.W1M_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.W1M_h4.setObjectName("W1M_h4")
        self.gridLayout_5.addWidget(self.W1M_h4, 17, 3, 1, 1)
        self.W1V_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.W1V_h4.setObjectName("W1V_h4")
        self.gridLayout_5.addWidget(self.W1V_h4, 17, 6, 1, 1)
        self.LGY1N_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1N_h3.setObjectName("LGY1N_h3")
        self.gridLayout_5.addWidget(self.LGY1N_h3, 13, 0, 1, 1)
        self.W1M_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.W1M_h3.setObjectName("W1M_h3")
        self.gridLayout_5.addWidget(self.W1M_h3, 13, 3, 1, 1)
        self.Y1L_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1L_h1.setObjectName("Y1L_h1")
        self.gridLayout_5.addWidget(self.Y1L_h1, 3, 3, 1, 1)
        self.LGY1K_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1K_h1.setObjectName("LGY1K_h1")
        self.gridLayout_5.addWidget(self.LGY1K_h1, 5, 1, 1, 1)
        self.W1R_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.W1R_h4.setObjectName("W1R_h4")
        self.gridLayout_5.addWidget(self.W1R_h4, 17, 5, 1, 1)
        self.C1A_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C1A_h4.setObjectName("C1A_h4")
        self.gridLayout_5.addWidget(self.C1A_h4, 15, 8, 1, 1)
        self.C1Y_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C1Y_h1.setObjectName("C1Y_h1")
        self.gridLayout_5.addWidget(self.C1Y_h1, 4, 6, 1, 1)
        self.Y1N_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1N_h1.setObjectName("Y1N_h1")
        self.gridLayout_5.addWidget(self.Y1N_h1, 3, 5, 1, 1)
        self.Y1K_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1K_h4.setObjectName("Y1K_h4")
        self.gridLayout_5.addWidget(self.Y1K_h4, 15, 1, 1, 1)
        self.C1A_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C1A_h1.setObjectName("C1A_h1")
        self.gridLayout_5.addWidget(self.C1A_h1, 3, 8, 1, 1)
        self.C2X_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.C2X_h1.setObjectName("C2X_h1")
        self.gridLayout_5.addWidget(self.C2X_h1, 4, 8, 1, 1)
        self.C1Y_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.C1Y_h4.setObjectName("C1Y_h4")
        self.gridLayout_5.addWidget(self.C1Y_h4, 16, 6, 1, 1)
        self.LGY1N_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1N_h1.setObjectName("LGY1N_h1")
        self.gridLayout_5.addWidget(self.LGY1N_h1, 5, 0, 1, 1)
        self.Y1P_h4 = QtWidgets.QCheckBox(self.widget_3)
        self.Y1P_h4.setObjectName("Y1P_h4")
        self.gridLayout_5.addWidget(self.Y1P_h4, 15, 6, 1, 1)
        self.h1 = QtWidgets.QCheckBox(self.widget_3)
        self.h1.setObjectName("h1")
        self.gridLayout_5.addWidget(self.h1, 1, 0, 1, 2)
        self.h3 = QtWidgets.QCheckBox(self.widget_3)
        self.h3.setObjectName("h3")
        self.gridLayout_5.addWidget(self.h3, 1, 6, 1, 3)
        self.Sunday = QtWidgets.QCheckBox(self.widget_3)
        self.Sunday.setObjectName("Sunday")
        self.gridLayout_5.addWidget(self.Sunday, 19, 9, 1, 1)
        self.h2 = QtWidgets.QCheckBox(self.widget_3)
        self.h2.setObjectName("h2")
        self.gridLayout_5.addWidget(self.h2, 1, 3, 1, 3)
        self.checkBox = QtWidgets.QCheckBox(self.widget_3)
        self.checkBox.setObjectName("checkBox")
        self.gridLayout_5.addWidget(self.checkBox, 20, 0, 1, 2)
        self.checkBox_2 = QtWidgets.QCheckBox(self.widget_3)
        self.checkBox_2.setObjectName("checkBox_2")
        self.gridLayout_5.addWidget(self.checkBox_2, 20, 3, 1, 2)
        self.LGY1L_h1 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1L_h1.setObjectName("LGY1L_h1")
        self.gridLayout_5.addWidget(self.LGY1L_h1, 5, 2, 1, 1)
        self.LGY1L_h2 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1L_h2.setObjectName("LGY1L_h2")
        self.gridLayout_5.addWidget(self.LGY1L_h2, 9, 2, 1, 1)
        self.LGY1L_h3 = QtWidgets.QCheckBox(self.widget_3)
        self.LGY1L_h3.setObjectName("LGY1L_h3")
        self.gridLayout_5.addWidget(self.LGY1L_h3, 13, 2, 1, 1)
        self.gridLayout_7.addWidget(self.widget_3, 0, 0, 1, 1)
        self.widget_4 = QtWidgets.QWidget(self.page_3)
        self.widget_4.setMinimumSize(QtCore.QSize(500, 0))
        self.widget_4.setStyleSheet("background-color: rgb(240, 240, 240);")
        self.widget_4.setObjectName("widget_4")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.widget_4)
        self.gridLayout_4.setContentsMargins(-1, -1, 11, -1)
        self.gridLayout_4.setVerticalSpacing(11)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.label_36 = QtWidgets.QLabel(self.widget_4)
        self.label_36.setObjectName("label_36")
        self.gridLayout_4.addWidget(self.label_36, 15, 0, 1, 2)
        self.prd_start_h3 = QtWidgets.QDateEdit(self.widget_4)
        self.prd_start_h3.setMaximumSize(QtCore.QSize(200, 16777215))
        self.prd_start_h3.setLocale(
            QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates)
        )
        self.prd_start_h3.setObjectName("prd_start_h3")
        self.gridLayout_4.addWidget(self.prd_start_h3, 9, 3, 1, 1)
        self.mutation_rate = QtWidgets.QLineEdit(self.widget_4)
        self.mutation_rate.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.mutation_rate.setFont(font)
        self.mutation_rate.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.mutation_rate.setAlignment(QtCore.Qt.AlignCenter)
        self.mutation_rate.setObjectName("mutation_rate")
        self.gridLayout_4.addWidget(self.mutation_rate, 15, 3, 1, 1)
        self.label_23 = QtWidgets.QLabel(self.widget_4)
        self.label_23.setObjectName("label_23")
        self.gridLayout_4.addWidget(self.label_23, 6, 1, 1, 1)
        self.man_power_h2 = QtWidgets.QLineEdit(self.widget_4)
        self.man_power_h2.setMaximumSize(QtCore.QSize(200, 16777215))
        self.man_power_h2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.man_power_h2.setAlignment(QtCore.Qt.AlignCenter)
        self.man_power_h2.setObjectName("man_power_h2")
        self.gridLayout_4.addWidget(self.man_power_h2, 4, 3, 1, 1)
        self.label_35 = QtWidgets.QLabel(self.widget_4)
        self.label_35.setObjectName("label_35")
        self.gridLayout_4.addWidget(self.label_35, 3, 1, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(
            40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum
        )
        self.gridLayout_4.addItem(spacerItem1, 6, 2, 1, 1)
        self.label_20 = QtWidgets.QLabel(self.widget_4)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_20.sizePolicy().hasHeightForWidth())
        self.label_20.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_20.setFont(font)
        self.label_20.setObjectName("label_20")
        self.gridLayout_4.addWidget(self.label_20, 0, 0, 1, 2)
        self.prd_start_h2 = QtWidgets.QDateEdit(self.widget_4)
        self.prd_start_h2.setMaximumSize(QtCore.QSize(200, 16777215))
        self.prd_start_h2.setLocale(
            QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates)
        )
        self.prd_start_h2.setObjectName("prd_start_h2")
        self.gridLayout_4.addWidget(self.prd_start_h2, 8, 3, 1, 1)
        self.label_39 = QtWidgets.QLabel(self.widget_4)
        self.label_39.setObjectName("label_39")
        self.gridLayout_4.addWidget(self.label_39, 9, 0, 1, 2)
        self.prd_start_h1 = QtWidgets.QDateEdit(self.widget_4)
        self.prd_start_h1.setMaximumSize(QtCore.QSize(200, 16777215))
        self.prd_start_h1.setLocale(
            QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates)
        )
        self.prd_start_h1.setMaximumDateTime(
            QtCore.QDateTime(QtCore.QDate(9999, 12, 31), QtCore.QTime(23, 59, 56))
        )
        self.prd_start_h1.setTimeSpec(QtCore.Qt.LocalTime)
        self.prd_start_h1.setObjectName("prd_start_h1")
        self.gridLayout_4.addWidget(self.prd_start_h1, 7, 3, 1, 1)
        self.man_power_h1 = QtWidgets.QLineEdit(self.widget_4)
        self.man_power_h1.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.man_power_h1.setFont(font)
        self.man_power_h1.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.man_power_h1.setAlignment(QtCore.Qt.AlignCenter)
        self.man_power_h1.setObjectName("man_power_h1")
        self.gridLayout_4.addWidget(self.man_power_h1, 3, 3, 1, 1)
        self.label_38 = QtWidgets.QLabel(self.widget_4)
        self.label_38.setObjectName("label_38")
        self.gridLayout_4.addWidget(self.label_38, 8, 0, 1, 2)
        self.label_37 = QtWidgets.QLabel(self.widget_4)
        self.label_37.setObjectName("label_37")
        self.gridLayout_4.addWidget(self.label_37, 7, 0, 1, 2)
        self.label_32 = QtWidgets.QLabel(self.widget_4)
        self.label_32.setObjectName("label_32")
        self.gridLayout_4.addWidget(self.label_32, 13, 0, 1, 2)
        self.label_33 = QtWidgets.QLabel(self.widget_4)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.MinimumExpanding
        )
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_33.sizePolicy().hasHeightForWidth())
        self.label_33.setSizePolicy(sizePolicy)
        self.label_33.setMaximumSize(QtCore.QSize(16777215, 100))
        self.label_33.setObjectName("label_33")
        self.gridLayout_4.addWidget(self.label_33, 2, 0, 1, 2)
        spacerItem2 = QtWidgets.QSpacerItem(
            80, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum
        )
        self.gridLayout_4.addItem(spacerItem2, 3, 2, 1, 1)
        self.population_size = QtWidgets.QLineEdit(self.widget_4)
        self.population_size.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.population_size.setFont(font)
        self.population_size.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.population_size.setAlignment(QtCore.Qt.AlignCenter)
        self.population_size.setObjectName("population_size")
        self.gridLayout_4.addWidget(self.population_size, 13, 3, 1, 1)
        self.generations = QtWidgets.QLineEdit(self.widget_4)
        self.generations.setMaximumSize(QtCore.QSize(200, 16777215))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.generations.setFont(font)
        self.generations.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.generations.setAlignment(QtCore.Qt.AlignCenter)
        self.generations.setObjectName("generations")
        self.gridLayout_4.addWidget(self.generations, 14, 3, 1, 1)
        self.label_40 = QtWidgets.QLabel(self.widget_4)
        self.label_40.setObjectName("label_40")
        self.gridLayout_4.addWidget(self.label_40, 10, 0, 1, 2)
        self.label_34 = QtWidgets.QLabel(self.widget_4)
        self.label_34.setObjectName("label_34")
        self.gridLayout_4.addWidget(self.label_34, 14, 0, 1, 2)
        spacerItem3 = QtWidgets.QSpacerItem(
            40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum
        )
        self.gridLayout_4.addItem(spacerItem3, 4, 2, 1, 1)
        self.prd_start_h4 = QtWidgets.QDateEdit(self.widget_4)
        self.prd_start_h4.setMaximumSize(QtCore.QSize(200, 16777215))
        self.prd_start_h4.setLocale(
            QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates)
        )
        self.prd_start_h4.setObjectName("prd_start_h4")
        self.gridLayout_4.addWidget(self.prd_start_h4, 10, 3, 1, 1)
        self.label_22 = QtWidgets.QLabel(self.widget_4)
        self.label_22.setMaximumSize(QtCore.QSize(16777215, 100))
        self.label_22.setObjectName("label_22")
        self.gridLayout_4.addWidget(self.label_22, 11, 0, 1, 2)
        spacerItem4 = QtWidgets.QSpacerItem(
            20, 0, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Maximum
        )
        self.gridLayout_4.addItem(spacerItem4, 1, 0, 1, 1)
        self.man_power_h3 = QtWidgets.QLineEdit(self.widget_4)
        self.man_power_h3.setMaximumSize(QtCore.QSize(200, 16777215))
        self.man_power_h3.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.man_power_h3.setAlignment(QtCore.Qt.AlignCenter)
        self.man_power_h3.setObjectName("man_power_h3")
        self.gridLayout_4.addWidget(self.man_power_h3, 5, 3, 1, 1)
        self.label_14 = QtWidgets.QLabel(self.widget_4)
        self.label_14.setObjectName("label_14")
        self.gridLayout_4.addWidget(self.label_14, 4, 1, 1, 1)
        self.man_power_h4 = QtWidgets.QLineEdit(self.widget_4)
        self.man_power_h4.setMaximumSize(QtCore.QSize(200, 16777215))
        self.man_power_h4.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.man_power_h4.setAlignment(QtCore.Qt.AlignCenter)
        self.man_power_h4.setObjectName("man_power_h4")
        self.gridLayout_4.addWidget(self.man_power_h4, 6, 3, 1, 1)
        self.label_30 = QtWidgets.QLabel(self.widget_4)
        self.label_30.setObjectName("label_30")
        self.gridLayout_4.addWidget(self.label_30, 16, 0, 1, 1)
        spacerItem5 = QtWidgets.QSpacerItem(
            40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum
        )
        self.gridLayout_4.addItem(spacerItem5, 5, 2, 1, 1)
        self.core_processing = QtWidgets.QLineEdit(self.widget_4)
        self.core_processing.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.core_processing.setAlignment(QtCore.Qt.AlignCenter)
        self.core_processing.setObjectName("core_processing")
        self.gridLayout_4.addWidget(self.core_processing, 16, 3, 1, 1)
        self.label_16 = QtWidgets.QLabel(self.widget_4)
        self.label_16.setObjectName("label_16")
        self.gridLayout_4.addWidget(self.label_16, 5, 1, 1, 1)
        self.population_line = QtWidgets.QLineEdit(self.widget_4)
        self.population_line.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.population_line.setAlignment(QtCore.Qt.AlignCenter)
        self.population_line.setObjectName("population_line")
        self.gridLayout_4.addWidget(self.population_line, 12, 3, 1, 1)
        self.label_31 = QtWidgets.QLabel(self.widget_4)
        self.label_31.setObjectName("label_31")
        self.gridLayout_4.addWidget(self.label_31, 12, 0, 1, 1)
        self.gridLayout_7.addWidget(self.widget_4, 0, 1, 1, 1)
        self.stackedWidget.addWidget(self.page_3)
        self.page_4 = QtWidgets.QWidget()
        self.page_4.setEnabled(True)
        self.page_4.setObjectName("page_4")
        self.Main_page = QtWidgets.QWidget(self.page_4)
        self.Main_page.setGeometry(QtCore.QRect(10, 10, 1133, 738))
        self.Main_page.setStyleSheet("background-color: rgb(240, 240, 240);")
        self.Main_page.setObjectName("Main_page")
        self.gridLayoutWidget = QtWidgets.QWidget(self.Main_page)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(90, 0, 170, 731))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.AP1222CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP1222CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP1222CW1W.setObjectName("AP1222CW1W")
        self.gridLayout.addWidget(self.AP1222CW1W, 7, 1, 1, 1)
        self.label_48 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_48.setObjectName("label_48")
        self.gridLayout.addWidget(self.label_48, 7, 0, 1, 1)
        self.label_49 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_49.setObjectName("label_49")
        self.gridLayout.addWidget(self.label_49, 8, 0, 1, 1)
        self.AP1223CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP1223CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP1223CW1W.setObjectName("AP1223CW1W")
        self.gridLayout.addWidget(self.AP1223CW1W, 8, 1, 1, 1)
        self.AP55023HR1GD = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP55023HR1GD.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP55023HR1GD.setObjectName("AP55023HR1GD")
        self.gridLayout.addWidget(self.AP55023HR1GD, 9, 1, 1, 1)
        self.label_50 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_50.setObjectName("label_50")
        self.gridLayout.addWidget(self.label_50, 9, 0, 1, 1)
        self.label_51 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_51.setObjectName("label_51")
        self.gridLayout.addWidget(self.label_51, 10, 0, 1, 1)
        self.AW1022CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AW1022CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1022CW1W.setObjectName("AW1022CW1W")
        self.gridLayout.addWidget(self.AW1022CW1W, 10, 1, 1, 1)
        self.AW1822CW3W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AW1822CW3W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1822CW3W.setObjectName("AW1822CW3W")
        self.gridLayout.addWidget(self.AW1822CW3W, 12, 1, 1, 1)
        self.label_52 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_52.setObjectName("label_52")
        self.gridLayout.addWidget(self.label_52, 11, 0, 1, 1)
        self.AW1023TW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AW1023TW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1023TW1W.setObjectName("AW1023TW1W")
        self.gridLayout.addWidget(self.AW1023TW1W, 11, 1, 1, 1)
        self.label_53 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_53.setObjectName("label_53")
        self.gridLayout.addWidget(self.label_53, 12, 0, 1, 1)
        self.label_54 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_54.setObjectName("label_54")
        self.gridLayout.addWidget(self.label_54, 13, 0, 1, 1)
        self.AW2422CW3W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AW2422CW3W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW2422CW3W.setObjectName("AW2422CW3W")
        self.gridLayout.addWidget(self.AW2422CW3W, 13, 1, 1, 1)
        self.label_55 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_55.setObjectName("label_55")
        self.gridLayout.addWidget(self.label_55, 14, 0, 1, 1)
        self.AW1822DR3W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AW1822DR3W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1822DR3W.setObjectName("AW1822DR3W")
        self.gridLayout.addWidget(self.AW1822DR3W, 14, 1, 1, 1)
        self.label_56 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_56.setObjectName("label_56")
        self.gridLayout.addWidget(self.label_56, 15, 0, 1, 1)
        self.AW1422CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AW1422CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1422CW1W.setObjectName("AW1422CW1W")
        self.gridLayout.addWidget(self.AW1422CW1W, 15, 1, 1, 1)
        self.AP0522CR1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP0522CR1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP0522CR1W.setObjectName("AP0522CR1W")
        self.gridLayout.addWidget(self.AP0522CR1W, 0, 1, 1, 1)
        self.label_42 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_42.setObjectName("label_42")
        self.gridLayout.addWidget(self.label_42, 1, 0, 1, 1)
        self.AP0621CR1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP0621CR1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP0621CR1W.setObjectName("AP0621CR1W")
        self.gridLayout.addWidget(self.AP0621CR1W, 1, 1, 1, 1)
        self.AP0722CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP0722CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP0722CW1W.setObjectName("AP0722CW1W")
        self.gridLayout.addWidget(self.AP0722CW1W, 2, 1, 1, 1)
        self.label_41 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_41.setObjectName("label_41")
        self.gridLayout.addWidget(self.label_41, 0, 0, 1, 1)
        self.label_43 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_43.setObjectName("label_43")
        self.gridLayout.addWidget(self.label_43, 2, 0, 1, 1)
        self.label_47 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_47.setObjectName("label_47")
        self.gridLayout.addWidget(self.label_47, 6, 0, 1, 1)
        self.label_44 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_44.setObjectName("label_44")
        self.gridLayout.addWidget(self.label_44, 3, 0, 1, 1)
        self.AP1022TW1GD = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP1022TW1GD.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP1022TW1GD.setObjectName("AP1022TW1GD")
        self.gridLayout.addWidget(self.AP1022TW1GD, 6, 1, 1, 1)
        self.AP1022CW1G = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP1022CW1G.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP1022CW1G.setObjectName("AP1022CW1G")
        self.gridLayout.addWidget(self.AP1022CW1G, 5, 1, 1, 1)
        self.AP0822CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP0822CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP0822CW1W.setObjectName("AP0822CW1W")
        self.gridLayout.addWidget(self.AP0822CW1W, 3, 1, 1, 1)
        self.AP1022HW1GD = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AP1022HW1GD.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AP1022HW1GD.setObjectName("AP1022HW1GD")
        self.gridLayout.addWidget(self.AP1022HW1GD, 4, 1, 1, 1)
        self.label_45 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_45.setObjectName("label_45")
        self.gridLayout.addWidget(self.label_45, 4, 0, 1, 1)
        self.label_46 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_46.setObjectName("label_46")
        self.gridLayout.addWidget(self.label_46, 5, 0, 1, 1)
        self.label_57 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_57.setObjectName("label_57")
        self.gridLayout.addWidget(self.label_57, 16, 0, 1, 1)
        self.AW1222CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget)
        self.AW1222CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1222CW1W.setObjectName("AW1222CW1W")
        self.gridLayout.addWidget(self.AW1222CW1W, 16, 1, 1, 1)
        self.gridLayoutWidget_2 = QtWidgets.QWidget(self.Main_page)
        self.gridLayoutWidget_2.setGeometry(QtCore.QRect(320, 0, 194, 731))
        self.gridLayoutWidget_2.setObjectName("gridLayoutWidget_2")
        self.gridLayout_9 = QtWidgets.QGridLayout(self.gridLayoutWidget_2)
        self.gridLayout_9.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_9.setObjectName("gridLayout_9")
        self.label_69 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_69.setObjectName("label_69")
        self.gridLayout_9.addWidget(self.label_69, 11, 0, 1, 1)
        self.LP0624WFR = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.LP0624WFR.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.LP0624WFR.setObjectName("LP0624WFR")
        self.gridLayout_9.addWidget(self.LP0624WFR, 11, 1, 1, 1)
        self.label_70 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_70.setObjectName("label_70")
        self.gridLayout_9.addWidget(self.label_70, 12, 0, 1, 1)
        self.AW1823TW3W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW1823TW3W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1823TW3W.setObjectName("AW1823TW3W")
        self.gridLayout_9.addWidget(self.AW1823TW3W, 2, 1, 1, 1)
        self.AW0822DR1W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW0822DR1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW0822DR1W.setObjectName("AW0822DR1W")
        self.gridLayout_9.addWidget(self.AW0822DR1W, 3, 1, 1, 1)
        self.label_61 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_61.setObjectName("label_61")
        self.gridLayout_9.addWidget(self.label_61, 3, 0, 1, 1)
        self.label_62 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_62.setObjectName("label_62")
        self.gridLayout_9.addWidget(self.label_62, 4, 0, 1, 1)
        self.AW1422TW1W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW1422TW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1422TW1W.setObjectName("AW1422TW1W")
        self.gridLayout_9.addWidget(self.AW1422TW1W, 4, 1, 1, 1)
        self.label_63 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_63.setObjectName("label_63")
        self.gridLayout_9.addWidget(self.label_63, 5, 0, 1, 1)
        self.AW2422DR3W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW2422DR3W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW2422DR3W.setObjectName("AW2422DR3W")
        self.gridLayout_9.addWidget(self.AW2422DR3W, 5, 1, 1, 1)
        self.AW0823TW1W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW0823TW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW0823TW1W.setObjectName("AW0823TW1W")
        self.gridLayout_9.addWidget(self.AW0823TW1W, 6, 1, 1, 1)
        self.label_65 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_65.setObjectName("label_65")
        self.gridLayout_9.addWidget(self.label_65, 7, 0, 1, 1)
        self.AW2223TW3W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW2223TW3W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW2223TW3W.setObjectName("AW2223TW3W")
        self.gridLayout_9.addWidget(self.AW2223TW3W, 8, 1, 1, 1)
        self.label_67 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_67.setObjectName("label_67")
        self.gridLayout_9.addWidget(self.label_67, 9, 0, 1, 1)
        self.label_64 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_64.setObjectName("label_64")
        self.gridLayout_9.addWidget(self.label_64, 6, 0, 1, 1)
        self.label_66 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_66.setObjectName("label_66")
        self.gridLayout_9.addWidget(self.label_66, 8, 0, 1, 1)
        self.label_68 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_68.setObjectName("label_68")
        self.gridLayout_9.addWidget(self.label_68, 10, 0, 1, 1)
        self.AW1221DR3W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW1221DR3W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1221DR3W.setObjectName("AW1221DR3W")
        self.gridLayout_9.addWidget(self.AW1221DR3W, 7, 1, 1, 1)
        self.LP0823GSSM = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.LP0823GSSM.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.LP0823GSSM.setObjectName("LP0823GSSM")
        self.gridLayout_9.addWidget(self.LP0823GSSM, 9, 1, 1, 1)
        self.LP0623WSR = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.LP0623WSR.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.LP0623WSR.setObjectName("LP0623WSR")
        self.gridLayout_9.addWidget(self.LP0623WSR, 10, 1, 1, 1)
        self.AW1222TW1W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW1222TW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW1222TW1W.setObjectName("AW1222TW1W")
        self.gridLayout_9.addWidget(self.AW1222TW1W, 0, 1, 1, 1)
        self.label_60 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_60.setObjectName("label_60")
        self.gridLayout_9.addWidget(self.label_60, 2, 0, 1, 1)
        self.AW0823CW1W = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AW0823CW1W.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AW0823CW1W.setObjectName("AW0823CW1W")
        self.gridLayout_9.addWidget(self.AW0823CW1W, 1, 1, 1, 1)
        self.label_58 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_58.setObjectName("label_58")
        self.gridLayout_9.addWidget(self.label_58, 0, 0, 1, 1)
        self.label_59 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_59.setObjectName("label_59")
        self.gridLayout_9.addWidget(self.label_59, 1, 0, 1, 1)
        self.AS_10TR4RYDTD02 = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AS_10TR4RYDTD02.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_10TR4RYDTD02.setObjectName("AS_10TR4RYDTD02")
        self.gridLayout_9.addWidget(self.AS_10TR4RYDTD02, 15, 1, 1, 1)
        self.label_71 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_71.setObjectName("label_71")
        self.gridLayout_9.addWidget(self.label_71, 13, 0, 1, 1)
        self.LP0524WFR = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.LP0524WFR.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.LP0524WFR.setObjectName("LP0524WFR")
        self.gridLayout_9.addWidget(self.LP0524WFR, 12, 1, 1, 1)
        self.LP0723WSR = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.LP0723WSR.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.LP0723WSR.setObjectName("LP0723WSR")
        self.gridLayout_9.addWidget(self.LP0723WSR, 13, 1, 1, 1)
        self.label_73 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_73.setObjectName("label_73")
        self.gridLayout_9.addWidget(self.label_73, 15, 0, 1, 1)
        self.label_72 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_72.setObjectName("label_72")
        self.gridLayout_9.addWidget(self.label_72, 14, 0, 1, 1)
        self.LP1023BSSM = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.LP1023BSSM.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.LP1023BSSM.setObjectName("LP1023BSSM")
        self.gridLayout_9.addWidget(self.LP1023BSSM, 14, 1, 1, 1)
        self.label_74 = QtWidgets.QLabel(self.gridLayoutWidget_2)
        self.label_74.setObjectName("label_74")
        self.gridLayout_9.addWidget(self.label_74, 16, 0, 1, 1)
        self.AS_10TR4RYD02D = QtWidgets.QLineEdit(self.gridLayoutWidget_2)
        self.AS_10TR4RYD02D.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_10TR4RYD02D.setObjectName("AS_10TR4RYD02D")
        self.gridLayout_9.addWidget(self.AS_10TR4RYD02D, 16, 1, 1, 1)
        self.gridLayoutWidget_3 = QtWidgets.QWidget(self.Main_page)
        self.gridLayoutWidget_3.setGeometry(QtCore.QRect(560, 0, 213, 731))
        self.gridLayoutWidget_3.setObjectName("gridLayoutWidget_3")
        self.gridLayout_11 = QtWidgets.QGridLayout(self.gridLayoutWidget_3)
        self.gridLayout_11.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_11.setObjectName("gridLayout_11")
        self.label_75 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_75.setObjectName("label_75")
        self.gridLayout_11.addWidget(self.label_75, 11, 0, 1, 1)
        self.AS_18TW4RGATD00O = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_18TW4RGATD00O.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_18TW4RGATD00O.setObjectName("AS_18TW4RGATD00O")
        self.gridLayout_11.addWidget(self.AS_18TW4RGATD00O, 11, 1, 1, 1)
        self.label_76 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_76.setObjectName("label_76")
        self.gridLayout_11.addWidget(self.label_76, 12, 0, 1, 1)
        self.AS_12TR4RYDTU00B = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12TR4RYDTU00B.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12TR4RYDTU00B.setObjectName("AS_12TR4RYDTU00B")
        self.gridLayout_11.addWidget(self.AS_12TR4RYDTU00B, 2, 1, 1, 1)
        self.AS_18TW4RGATU00 = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_18TW4RGATU00.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_18TW4RGATU00.setObjectName("AS_18TW4RGATU00")
        self.gridLayout_11.addWidget(self.AS_18TW4RGATU00, 3, 1, 1, 1)
        self.label_77 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_77.setObjectName("label_77")
        self.gridLayout_11.addWidget(self.label_77, 3, 0, 1, 1)
        self.label_78 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_78.setObjectName("label_78")
        self.gridLayout_11.addWidget(self.label_78, 4, 0, 1, 1)
        self.AS_12CR4RVETD01 = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12CR4RVETD01.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12CR4RVETD01.setObjectName("AS_12CR4RVETD01")
        self.gridLayout_11.addWidget(self.AS_12CR4RVETD01, 4, 1, 1, 1)
        self.label_79 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_79.setObjectName("label_79")
        self.gridLayout_11.addWidget(self.label_79, 5, 0, 1, 1)
        self.AS_18TW4RGATD00 = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_18TW4RGATD00.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_18TW4RGATD00.setObjectName("AS_18TW4RGATD00")
        self.gridLayout_11.addWidget(self.AS_18TW4RGATD00, 5, 1, 1, 1)
        self.AS_12TR4RYDTD00B = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12TR4RYDTD00B.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12TR4RYDTD00B.setObjectName("AS_12TR4RYDTD00B")
        self.gridLayout_11.addWidget(self.AS_12TR4RYDTD00B, 6, 1, 1, 1)
        self.label_80 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_80.setObjectName("label_80")
        self.gridLayout_11.addWidget(self.label_80, 7, 0, 1, 1)
        self.AS_12CR4RVE01D = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12CR4RVE01D.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12CR4RVE01D.setObjectName("AS_12CR4RVE01D")
        self.gridLayout_11.addWidget(self.AS_12CR4RVE01D, 8, 1, 1, 1)
        self.label_81 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_81.setObjectName("label_81")
        self.gridLayout_11.addWidget(self.label_81, 9, 0, 1, 1)
        self.label_82 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_82.setObjectName("label_82")
        self.gridLayout_11.addWidget(self.label_82, 6, 0, 1, 1)
        self.label_83 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_83.setObjectName("label_83")
        self.gridLayout_11.addWidget(self.label_83, 8, 0, 1, 1)
        self.label_84 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_84.setObjectName("label_84")
        self.gridLayout_11.addWidget(self.label_84, 10, 0, 1, 1)
        self.AS_12TR4RYD00BD = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12TR4RYD00BD.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12TR4RYD00BD.setObjectName("AS_12TR4RYD00BD")
        self.gridLayout_11.addWidget(self.AS_12TR4RYD00BD, 7, 1, 1, 1)
        self.AS_18TW4RGA00D = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_18TW4RGA00D.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_18TW4RGA00D.setObjectName("AS_18TW4RGA00D")
        self.gridLayout_11.addWidget(self.AS_18TW4RGA00D, 9, 1, 1, 1)
        self.HAP0824TWD = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.HAP0824TWD.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.HAP0824TWD.setObjectName("HAP0824TWD")
        self.gridLayout_11.addWidget(self.HAP0824TWD, 10, 1, 1, 1)
        self.AS_10TR4RYDTU02 = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_10TR4RYDTU02.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_10TR4RYDTU02.setObjectName("AS_10TR4RYDTU02")
        self.gridLayout_11.addWidget(self.AS_10TR4RYDTU02, 0, 1, 1, 1)
        self.label_85 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_85.setObjectName("label_85")
        self.gridLayout_11.addWidget(self.label_85, 2, 0, 1, 1)
        self.AS_12CR4RVEDJ01 = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12CR4RVEDJ01.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12CR4RVEDJ01.setObjectName("AS_12CR4RVEDJ01")
        self.gridLayout_11.addWidget(self.AS_12CR4RVEDJ01, 1, 1, 1, 1)
        self.label_86 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_86.setObjectName("label_86")
        self.gridLayout_11.addWidget(self.label_86, 0, 0, 1, 1)
        self.label_87 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_87.setObjectName("label_87")
        self.gridLayout_11.addWidget(self.label_87, 1, 0, 1, 1)
        self.AS_10TR4RYDTD02O = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_10TR4RYDTD02O.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_10TR4RYDTD02O.setObjectName("AS_10TR4RYDTD02O")
        self.gridLayout_11.addWidget(self.AS_10TR4RYDTD02O, 15, 1, 1, 1)
        self.label_88 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_88.setObjectName("label_88")
        self.gridLayout_11.addWidget(self.label_88, 13, 0, 1, 1)
        self.AS_12TR4RYDTD00BO = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12TR4RYDTD00BO.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12TR4RYDTD00BO.setObjectName("AS_12TR4RYDTD00BO")
        self.gridLayout_11.addWidget(self.AS_12TR4RYDTD00BO, 12, 1, 1, 1)
        self.AS_12CR4RVETD01O = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_12CR4RVETD01O.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_12CR4RVETD01O.setObjectName("AS_12CR4RVETD01O")
        self.gridLayout_11.addWidget(self.AS_12CR4RVETD01O, 13, 1, 1, 1)
        self.label_89 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_89.setObjectName("label_89")
        self.gridLayout_11.addWidget(self.label_89, 15, 0, 1, 1)
        self.label_90 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_90.setObjectName("label_90")
        self.gridLayout_11.addWidget(self.label_90, 14, 0, 1, 1)
        self.AS_10CR4RYDTD02O = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.AS_10CR4RYDTD02O.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.AS_10CR4RYDTD02O.setObjectName("AS_10CR4RYDTD02O")
        self.gridLayout_11.addWidget(self.AS_10CR4RYDTD02O, 14, 1, 1, 1)
        self.label_91 = QtWidgets.QLabel(self.gridLayoutWidget_3)
        self.label_91.setObjectName("label_91")
        self.gridLayout_11.addWidget(self.label_91, 16, 0, 1, 1)
        self.RMV08A10A = QtWidgets.QLineEdit(self.gridLayoutWidget_3)
        self.RMV08A10A.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.RMV08A10A.setObjectName("RMV08A10A")
        self.gridLayout_11.addWidget(self.RMV08A10A, 16, 1, 1, 1)
        self.label_103 = QtWidgets.QLabel(self.Main_page)
        self.label_103.setGeometry(QtCore.QRect(850, 20, 79, 22))
        self.label_103.setObjectName("label_103")
        self.label_104 = QtWidgets.QLabel(self.Main_page)
        self.label_104.setGeometry(QtCore.QRect(850, 60, 79, 22))
        self.label_104.setObjectName("label_104")
        self.RMV10A10A = QtWidgets.QLineEdit(self.Main_page)
        self.RMV10A10A.setGeometry(QtCore.QRect(950, 20, 63, 22))
        self.RMV10A10A.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.RMV10A10A.setObjectName("RMV10A10A")
        self.RMV12A10A = QtWidgets.QLineEdit(self.Main_page)
        self.RMV12A10A.setGeometry(QtCore.QRect(950, 60, 63, 22))
        self.RMV12A10A.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.RMV12A10A.setObjectName("RMV12A10A")
        self.stackedWidget.addWidget(self.page_4)
        self.gridLayout_10.addWidget(self.stackedWidget, 2, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.stackedWidget.setCurrentIndex(1)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "OEM Auto Planning"))
        self.Home.setText(_translate("MainWindow", "Home"))
        self.config.setText(_translate("MainWindow", "Configuration"))
        self.uph.setText(_translate("MainWindow", "UPH"))
        self.label_5.setText(_translate("MainWindow", "DATA PLANNING :"))
        self.label_6.setText(_translate("MainWindow", "    SHEET NAME :"))
        self.label_2.setText(_translate("MainWindow", "PRODUCTION PLAN"))
        self.label_3.setText(_translate("MainWindow", "ORDER PLAN"))
        self.label_8.setText(_translate("MainWindow", "    SHEET NAME :"))
        self.pushButton_3.setText(_translate("MainWindow", "Brown"))
        self.pushButton_2.setText(_translate("MainWindow", "Brown"))
        self.pushButton.setText(_translate("MainWindow", "Brown"))
        self.label_4.setText(_translate("MainWindow", "SALE PLAN"))
        self.label_7.setText(_translate("MainWindow", "    SHEET NAME :"))
        self.pushButton_4.setText(_translate("MainWindow", "Run"))
        self.pushButton_5.setText(_translate("MainWindow", "Clear Data"))
        self.label_17.setText(_translate("MainWindow", "SUPPLIER CONSTRAIN FOUL:"))
        self.label_10.setText(_translate("MainWindow", "TIME"))
        self.label_11.setText(_translate("MainWindow", "STATUS :"))
        self.label_15.setText(_translate("MainWindow", "FOLLOW THE SALES PLAN :"))
        self.label_9.setText(_translate("MainWindow", "SET UP TIME :"))
        self.label_13.setText(_translate("MainWindow", "WAITING TIME :"))
        self.label_19.setText(_translate("MainWindow", "NUMBER OF DAY TO OPEN OT :"))
        self.label_12.setText(_translate("MainWindow", "DAY"))
        self.label_18.setText(_translate("MainWindow", "UNIT"))
        self.status.setText(_translate("MainWindow", "No Plan..."))
        self.C1U_h3.setText(_translate("MainWindow", "C1U"))
        self.C3D_h4.setText(_translate("MainWindow", "C3D"))
        self.Y1N_h3.setText(_translate("MainWindow", "Y1N"))
        self.C1W_h3.setText(_translate("MainWindow", "C1W"))
        self.C1V_h3.setText(_translate("MainWindow", "C1V"))
        self.Y1P_h3.setText(_translate("MainWindow", "Y1P"))
        self.Friday.setText(_translate("MainWindow", "Friday"))
        self.LGY1K_h3.setText(_translate("MainWindow", "LG-Y1K"))
        self.W1V_h1.setText(_translate("MainWindow", "W1V"))
        self.Saturday.setText(_translate("MainWindow", "Saturday"))
        self.C2X_h2.setText(_translate("MainWindow", "C2X"))
        self.C1B_h2.setText(_translate("MainWindow", "C1B"))
        self.C3D_h2.setText(_translate("MainWindow", "C3D"))
        self.C3D_h3.setText(_translate("MainWindow", "C3D"))
        self.C2X_h3.setText(_translate("MainWindow", "C2X"))
        self.C1B_h3.setText(_translate("MainWindow", "C1B"))
        self.C1B_h4.setText(_translate("MainWindow", "C1B"))
        self.C1A_h3.setText(_translate("MainWindow", "C1A"))
        self.C1A_h2.setText(_translate("MainWindow", "C1A"))
        self.Y1L_h3.setText(_translate("MainWindow", "Y1L"))
        self.C1Y_h3.setText(_translate("MainWindow", "C1Y"))
        self.Tuesday.setText(_translate("MainWindow", "Tuesday"))
        self.W1R_h3.setText(_translate("MainWindow", "W1R"))
        self.W1M_h1.setText(_translate("MainWindow", "W1M"))
        self.Wednesday.setText(_translate("MainWindow", "Wednesday"))
        self.W1R_h1.setText(_translate("MainWindow", "W1R"))
        self.label_26.setText(_translate("MainWindow", "MODEL HISENSE 1 :"))
        self.Thursday.setText(_translate("MainWindow", "Thursday"))
        self.Y1G_h2.setText(_translate("MainWindow", "Y1G"))
        self.Y1K_h1.setText(_translate("MainWindow", "Y1K"))
        self.label_25.setText(_translate("MainWindow", "PRODUCTION LINE :"))
        self.Y1G_h1.setText(_translate("MainWindow", "Y1G"))
        self.Y1G_h3.setText(_translate("MainWindow", "Y1G"))
        self.C1F_h1.setText(_translate("MainWindow", "C1F"))
        self.Y1K_h3.setText(_translate("MainWindow", "Y1K"))
        self.Y1K_h2.setText(_translate("MainWindow", "Y1K"))
        self.Y1L_h2.setText(_translate("MainWindow", "Y1L"))
        self.Y1N_h2.setText(_translate("MainWindow", "Y1N"))
        self.label_27.setText(_translate("MainWindow", "MODEL HISENSE 2 :"))
        self.Y1G_h4.setText(_translate("MainWindow", "Y1G"))
        self.h4.setText(_translate("MainWindow", "HISENSE 4"))
        self.C1V_h2.setText(_translate("MainWindow", "C1V"))
        self.C1W_h2.setText(_translate("MainWindow", "C1W"))
        self.Y1P_h2.setText(_translate("MainWindow", "Y1P"))
        self.label_21.setText(_translate("MainWindow", "OT PLAN :"))
        self.label_28.setText(_translate("MainWindow", "MODEL HISENSE 3 :"))
        self.label_24.setText(_translate("MainWindow", "PLAN TYPE :"))
        self.C1F_h4.setText(_translate("MainWindow", "C1F"))
        self.C1W_h1.setText(_translate("MainWindow", "C1W"))
        self.Monday.setText(_translate("MainWindow", "Monday"))
        self.new_season.setText(_translate("MainWindow", "NEW SEASON"))
        self.W1V_h2.setText(_translate("MainWindow", "W1V"))
        self.W1M_h2.setText(_translate("MainWindow", "W1M"))
        self.Y1N_h4.setText(_translate("MainWindow", "Y1N"))
        self.W1V_h3.setText(_translate("MainWindow", "W1V"))
        self.C1Y_h2.setText(_translate("MainWindow", "C1Y"))
        self.C1F_h3.setText(_translate("MainWindow", "C1F"))
        self.C1F_h2.setText(_translate("MainWindow", "C1F"))
        self.C1U_h4.setText(_translate("MainWindow", "C1U"))
        self.Y1L_h4.setText(_translate("MainWindow", "Y1L"))
        self.C1U_h2.setText(_translate("MainWindow", "C1U"))
        self.LGY1K_h2.setText(_translate("MainWindow", "LG-Y1K"))
        self.C1V_h1.setText(_translate("MainWindow", "C1V"))
        self.C1V_h4.setText(_translate("MainWindow", "C1V"))
        self.continue_plan.setText(_translate("MainWindow", "CONTINUE PLAN"))
        self.label_29.setText(_translate("MainWindow", "MODEL HISENSE 4 :"))
        self.C3D_h1.setText(_translate("MainWindow", "C3D"))
        self.C1B_h1.setText(_translate("MainWindow", "C1B"))
        self.C2X_h4.setText(_translate("MainWindow", "C2X"))
        self.Y1P_h1.setText(_translate("MainWindow", "Y1P"))
        self.LGY1N_h2.setText(_translate("MainWindow", "LG-Y1N"))
        self.LGY1K_h4.setText(_translate("MainWindow", "LG-Y1K"))
        self.C1W_h4.setText(_translate("MainWindow", "C1W"))
        self.C1U_h1.setText(_translate("MainWindow", "C1U"))
        self.W1R_h2.setText(_translate("MainWindow", "W1R"))
        self.LGY1N_h4.setText(_translate("MainWindow", "LG-Y1N"))
        self.W1M_h4.setText(_translate("MainWindow", "W1M"))
        self.W1V_h4.setText(_translate("MainWindow", "W1V"))
        self.LGY1N_h3.setText(_translate("MainWindow", "LG-Y1N"))
        self.W1M_h3.setText(_translate("MainWindow", "W1M"))
        self.Y1L_h1.setText(_translate("MainWindow", "Y1L"))
        self.LGY1K_h1.setText(_translate("MainWindow", "LG-Y1K"))
        self.W1R_h4.setText(_translate("MainWindow", "W1R"))
        self.C1A_h4.setText(_translate("MainWindow", "C1A"))
        self.C1Y_h1.setText(_translate("MainWindow", "C1Y"))
        self.Y1N_h1.setText(_translate("MainWindow", "Y1N"))
        self.Y1K_h4.setText(_translate("MainWindow", "Y1K"))
        self.C1A_h1.setText(_translate("MainWindow", "C1A"))
        self.C2X_h1.setText(_translate("MainWindow", "C2X"))
        self.C1Y_h4.setText(_translate("MainWindow", "C1Y"))
        self.LGY1N_h1.setText(_translate("MainWindow", "LG-Y1N"))
        self.Y1P_h4.setText(_translate("MainWindow", "Y1P"))
        self.h1.setText(_translate("MainWindow", "HISENSE 1"))
        self.h3.setText(_translate("MainWindow", "HISENSE 3"))
        self.Sunday.setText(_translate("MainWindow", "Sunday"))
        self.h2.setText(_translate("MainWindow", "HISENSE 2"))
        self.checkBox.setText(_translate("MainWindow", "2.5 Hours"))
        self.checkBox_2.setText(_translate("MainWindow", "4 Hours"))
        self.LGY1L_h1.setText(_translate("MainWindow", "LG-Y1L"))
        self.LGY1L_h2.setText(_translate("MainWindow", "LG-Y1L"))
        self.LGY1L_h3.setText(_translate("MainWindow", "LG-Y1L"))
        self.label_36.setText(_translate("MainWindow", "MUTATION RATE :"))
        self.label_23.setText(_translate("MainWindow", "MAN POWER H4 :"))
        self.label_35.setText(_translate("MainWindow", "MAN POWER  H1 :"))
        self.label_20.setText(_translate("MainWindow", "PARAMETER :"))
        self.label_39.setText(_translate("MainWindow", "PRODUCTION START H3 :"))
        self.label_38.setText(_translate("MainWindow", "PRODUCTION START H2 :"))
        self.label_37.setText(_translate("MainWindow", "PRODUCTION START H1 :"))
        self.label_32.setText(_translate("MainWindow", "POPULATION SIZE :"))
        self.label_33.setText(_translate("MainWindow", "PRODUCTION PLAN "))
        self.label_40.setText(_translate("MainWindow", "PRODUCTION START H4 :"))
        self.label_34.setText(_translate("MainWindow", "GENERATION :"))
        self.label_22.setText(_translate("MainWindow", "GENETIC ALGORITHM "))
        self.label_14.setText(_translate("MainWindow", "MAN POWER H2 :"))
        self.label_30.setText(_translate("MainWindow", "Core Processing : "))
        self.label_16.setText(_translate("MainWindow", "MAN POWER H3 :"))
        self.label_31.setText(_translate("MainWindow", "POPULATION LINE :"))
        self.label_48.setText(_translate("MainWindow", "AP1222CW1W :"))
        self.label_49.setText(_translate("MainWindow", "AP1223CW1W :"))
        self.label_50.setText(_translate("MainWindow", "AP55023HR1GD :"))
        self.label_51.setText(_translate("MainWindow", "AW1022CW1W :"))
        self.label_52.setText(_translate("MainWindow", "AW1023TW1W :"))
        self.label_53.setText(_translate("MainWindow", "AW1822CW3W :"))
        self.label_54.setText(_translate("MainWindow", "AW2422CW3W :"))
        self.label_55.setText(_translate("MainWindow", "AW1822DR3W :"))
        self.label_56.setText(_translate("MainWindow", "AW1422CW1W :"))
        self.label_42.setText(_translate("MainWindow", "AP0621CR1W :"))
        self.label_41.setText(_translate("MainWindow", "AP0522CR1W :"))
        self.label_43.setText(_translate("MainWindow", "AP0722CW1W :"))
        self.label_47.setText(_translate("MainWindow", "AP1022TW1GD :"))
        self.label_44.setText(_translate("MainWindow", "AP0822CW1W :"))
        self.label_45.setText(_translate("MainWindow", "AP1022HW1GD :"))
        self.label_46.setText(_translate("MainWindow", "AP1022CW1G :"))
        self.label_57.setText(_translate("MainWindow", "AW1222CW1W :"))
        self.label_69.setText(_translate("MainWindow", "LP0624WFR :"))
        self.label_70.setText(_translate("MainWindow", "LP0524WFR :"))
        self.label_61.setText(_translate("MainWindow", "AW0822DR1W :"))
        self.label_62.setText(_translate("MainWindow", "AW1422TW1W :"))
        self.label_63.setText(_translate("MainWindow", "AW2422DR3W :"))
        self.label_65.setText(_translate("MainWindow", "AW1221DR3W :"))
        self.label_67.setText(_translate("MainWindow", "LP0823GSSM :"))
        self.label_64.setText(_translate("MainWindow", "AW0823TW1W :"))
        self.label_66.setText(_translate("MainWindow", "AW2223TW3W :"))
        self.label_68.setText(_translate("MainWindow", "LP0623WSR :"))
        self.label_60.setText(_translate("MainWindow", "AW1823TW3W :"))
        self.label_58.setText(_translate("MainWindow", "AW1222TW1W :"))
        self.label_59.setText(_translate("MainWindow", "AW0823CW1W :"))
        self.label_71.setText(_translate("MainWindow", "LP0723WSR :"))
        self.label_73.setText(_translate("MainWindow", "AS-10TR4RYDTD02 :"))
        self.label_72.setText(_translate("MainWindow", "LP1023BSSM :"))
        self.label_74.setText(_translate("MainWindow", "AS-10TR4RYD02(D) :"))
        self.label_75.setText(_translate("MainWindow", "AS-18TW4RGATD00-O :"))
        self.label_76.setText(_translate("MainWindow", "AS-12TR4RYDTD00B-O :"))
        self.label_77.setText(_translate("MainWindow", "AS-18TW4RGATU00 :"))
        self.label_78.setText(_translate("MainWindow", "AS-12CR4RVETD01 :"))
        self.label_79.setText(_translate("MainWindow", "AS-18TW4RGATD00 :"))
        self.label_80.setText(_translate("MainWindow", "AS-12TR4RYD00B(D) :"))
        self.label_81.setText(_translate("MainWindow", "AS-18TW4RGA00(D) :"))
        self.label_82.setText(_translate("MainWindow", "AS-12TR4RYDTD00B :"))
        self.label_83.setText(_translate("MainWindow", "AS-12CR4RVE01(D) :"))
        self.label_84.setText(_translate("MainWindow", "HAP0824TWD :"))
        self.label_85.setText(_translate("MainWindow", "AS-12TR4RYDTU00B :"))
        self.label_86.setText(_translate("MainWindow", "AS-10TR4RYDTU02 :"))
        self.label_87.setText(_translate("MainWindow", "AS-12CR4RVEDJ01 :"))
        self.label_88.setText(_translate("MainWindow", "AS-12CR4RVETD01-O :"))
        self.label_89.setText(_translate("MainWindow", "AS-10TR4RYDTD02-O :"))
        self.label_90.setText(_translate("MainWindow", "AS-10CR4RYDTD02-O :"))
        self.label_91.setText(_translate("MainWindow", "RMV08A10A :"))
        self.label_103.setText(_translate("MainWindow", "RMV10A10A :"))
        self.label_104.setText(_translate("MainWindow", "RMV12A10A :"))

        self.Home.clicked.connect(self.on_Home_toggled)
        self.config.clicked.connect(self.on_Config_toggled)
        self.uph.clicked.connect(self.on_uph_toggled)
        self.h1.clicked.connect(lambda: self.on_click_prd_line(self.h1))
        self.h2.clicked.connect(lambda: self.on_click_prd_line(self.h2))
        self.h3.clicked.connect(lambda: self.on_click_prd_line(self.h3))
        self.h4.clicked.connect(lambda: self.on_click_prd_line(self.h4))

        self.C1A_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C1A_h1))
        self.C1B_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C1B_h1))
        self.C1F_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C1F_h1))
        self.C1V_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C1V_h1))
        self.C1U_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C1U_h1))
        self.C1W_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C1W_h1))
        self.C1Y_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C1Y_h1))
        self.C2X_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C2X_h1))
        self.C3D_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.C3D_h1))
        self.Y1G_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.Y1G_h1))
        self.Y1K_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.Y1K_h1))
        self.Y1L_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.Y1L_h1))
        self.Y1N_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.Y1N_h1))
        self.Y1P_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.Y1P_h1))
        self.LGY1L_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.LGY1L_h1))
        self.LGY1N_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.LGY1N_h1))
        self.LGY1K_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.LGY1K_h1))
        self.W1M_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.W1M_h1))
        self.W1R_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.W1R_h1))
        self.W1V_h1.clicked.connect(lambda: self.on_click_checkbox_h1(self.W1V_h1))

        self.C1A_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C1A_h2))
        self.C1B_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C1B_h2))
        self.C1F_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C1F_h2))
        self.C1V_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C1V_h2))
        self.C1U_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C1U_h2))
        self.C1W_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C1W_h2))
        self.C1Y_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C1Y_h2))
        self.C2X_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C2X_h2))
        self.C3D_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.C3D_h2))
        self.Y1G_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.Y1G_h2))
        self.Y1K_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.Y1K_h2))
        self.Y1L_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.Y1L_h2))
        self.Y1N_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.Y1N_h2))
        self.Y1P_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.Y1P_h2))
        self.LGY1L_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.LGY1L_h2))
        self.LGY1N_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.LGY1N_h2))
        self.LGY1K_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.LGY1K_h2))
        self.W1M_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.W1M_h2))
        self.W1R_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.W1R_h2))
        self.W1V_h2.clicked.connect(lambda: self.on_click_checkbox_h2(self.W1V_h2))

        self.C1A_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C1A_h3))
        self.C1B_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C1B_h3))
        self.C1F_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C1F_h3))
        self.C1V_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C1V_h3))
        self.C1U_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C1U_h3))
        self.C1W_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C1W_h3))
        self.C1Y_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C1Y_h3))
        self.C2X_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C2X_h3))
        self.C3D_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.C3D_h3))
        self.Y1G_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.Y1G_h3))
        self.Y1K_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.Y1K_h3))
        self.Y1L_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.Y1L_h3))
        self.Y1N_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.Y1N_h3))
        self.Y1P_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.Y1P_h3))
        self.LGY1L_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.LGY1L_h3))
        self.LGY1N_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.LGY1N_h3))
        self.LGY1K_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.LGY1K_h3))
        self.W1M_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.W1M_h3))
        self.W1R_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.W1R_h3))
        self.W1V_h3.clicked.connect(lambda: self.on_click_checkbox_h3(self.W1V_h3))

        self.C1A_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C1A_h4))
        self.C1B_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C1B_h4))
        self.C1F_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C1F_h4))
        self.C1V_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C1V_h4))
        self.C1U_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C1U_h4))
        self.C1W_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C1W_h4))
        self.C1Y_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C1Y_h4))
        self.C2X_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C2X_h4))
        self.C3D_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.C3D_h4))
        self.Y1G_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.Y1G_h4))
        self.Y1K_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.Y1K_h4))
        self.Y1L_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.Y1L_h4))
        self.Y1N_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.Y1N_h4))
        self.Y1P_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.Y1P_h4))
        self.LGY1N_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.LGY1N_h4))
        self.LGY1K_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.LGY1K_h4))
        self.W1M_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.W1M_h4))
        self.W1R_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.W1R_h4))
        self.W1V_h4.clicked.connect(lambda: self.on_click_checkbox_h4(self.W1V_h4))

        self.new_season.clicked.connect(lambda: self.status_plan(self.new_season))
        self.continue_plan.clicked.connect(lambda: self.status_plan(self.continue_plan))

        self.Monday.clicked.connect(lambda: self.ot_day(self.Monday))
        self.Tuesday.clicked.connect(lambda: self.ot_day(self.Tuesday))
        self.Wednesday.clicked.connect(lambda: self.ot_day(self.Wednesday))
        self.Thursday.clicked.connect(lambda: self.ot_day(self.Thursday))
        self.Friday.clicked.connect(lambda: self.ot_day(self.Friday))
        self.Saturday.clicked.connect(lambda: self.ot_day(self.Saturday))
        self.Sunday.clicked.connect(lambda: self.ot_day(self.Sunday))

        self.prd_start_h1.dateChanged.connect(
            lambda date: self.onDateChanged("prd_start_h1", date)
        )
        self.prd_start_h2.dateChanged.connect(
            lambda date: self.onDateChanged("prd_start_h2", date)
        )
        self.prd_start_h3.dateChanged.connect(
            lambda date: self.onDateChanged("prd_start_h3", date)
        )
        self.prd_start_h4.dateChanged.connect(
            lambda date: self.onDateChanged("prd_start_h4", date)
        )

        # Genetic Parameter
        self.man_power_h1.textChanged.connect(self.onTextChanged("man_power_h1"))
        self.man_power_h2.textChanged.connect(self.onTextChanged("man_power_h2"))
        self.man_power_h3.textChanged.connect(self.onTextChanged("man_power_h3"))
        self.man_power_h4.textChanged.connect(self.onTextChanged("man_power_h4"))
        self.core_processing.textChanged.connect(self.onTextChanged("core_processing"))
        self.population_line.textChanged.connect(self.onTextChanged("population_line"))
        self.population_size.textChanged.connect(self.onTextChanged("population_size"))
        self.generations.textChanged.connect(self.onTextChanged("generations"))
        self.mutation_rate.textChanged.connect(self.onTextChanged("mutation_rate"))
        self.file_path.textChanged.connect(self.onTextChanged("file_path"))
        self.order_path.textChanged.connect(self.onTextChanged("order_path"))
        self.sale_path.textChanged.connect(self.onTextChanged("sale_path"))
        self.prd_sheet_name.textChanged.connect(self.onTextChanged("prd_sheet_name"))
        self.order_sheet_name.textChanged.connect(
            self.onTextChanged("order_sheet_name")
        )
        self.sale_sheet_name.textChanged.connect(self.onTextChanged("sale_sheet_name"))

        # UPH
        self.AP0522CR1W.textChanged.connect(self.onTextChangedUPH("AP0522CR1W"))
        self.AP0621CR1W.textChanged.connect(self.onTextChangedUPH("AP0621CR1W"))
        self.AP0722CW1W.textChanged.connect(self.onTextChangedUPH("AP0722CW1W"))
        self.AP0822CW1W.textChanged.connect(self.onTextChangedUPH("AP0822CW1W"))
        self.AP1022HW1GD.textChanged.connect(self.onTextChangedUPH("AP1022HW1GD"))
        self.AP1022CW1G.textChanged.connect(self.onTextChangedUPH("AP1022CW1G"))
        self.AP1022TW1GD.textChanged.connect(self.onTextChangedUPH("AP1022TW1GD"))
        self.AP1222CW1W.textChanged.connect(self.onTextChangedUPH("AP1222CW1W"))
        self.AP1223CW1W.textChanged.connect(self.onTextChangedUPH("AP1223CW1W"))
        self.AP55023HR1GD.textChanged.connect(self.onTextChangedUPH("AP55023HR1GD"))
        self.AW1022CW1W.textChanged.connect(self.onTextChangedUPH("AW1022CW1W"))
        self.AW1023TW1W.textChanged.connect(self.onTextChangedUPH("AW1023TW1W"))
        self.AW1822CW3W.textChanged.connect(self.onTextChangedUPH("AW1822CW3W"))
        self.AW2422CW3W.textChanged.connect(self.onTextChangedUPH("AW2422CW3W"))
        self.AW1822DR3W.textChanged.connect(self.onTextChangedUPH("AW1822DR3W"))
        self.AW1422CW1W.textChanged.connect(self.onTextChangedUPH("AW1422CW1W"))
        self.AW1222CW1W.textChanged.connect(self.onTextChangedUPH("AW1222CW1W"))

        self.AW1222TW1W.textChanged.connect(self.onTextChangedUPH("AW1222TW1W"))
        self.AW0822DR1W.textChanged.connect(self.onTextChangedUPH("AW0822DR1W"))
        self.AW0823CW1W.textChanged.connect(self.onTextChangedUPH("AW0823CW1W"))
        self.AW1823TW3W.textChanged.connect(self.onTextChangedUPH("AW1823TW3W"))
        self.AW1422TW1W.textChanged.connect(self.onTextChangedUPH("AW1422TW1W"))
        self.AW2422DR3W.textChanged.connect(self.onTextChangedUPH("AW2422DR3W"))
        self.AW0823TW1W.textChanged.connect(self.onTextChangedUPH("AW0823TW1W"))
        self.AW1221DR3W.textChanged.connect(self.onTextChangedUPH("AW1221DR3W"))
        self.AW2223TW3W.textChanged.connect(self.onTextChangedUPH("AW2223TW3W"))
        self.LP0823GSSM.textChanged.connect(self.onTextChangedUPH("LP0823GSSM"))
        self.LP0623WSR.textChanged.connect(self.onTextChangedUPH("LP0623WSR"))
        self.LP0624WFR.textChanged.connect(self.onTextChangedUPH("LP0624WFR"))
        self.LP0524WFR.textChanged.connect(self.onTextChangedUPH("LP0524WFR"))
        self.LP0723WSR.textChanged.connect(self.onTextChangedUPH("LP0723WSR"))
        self.LP1023BSSM.textChanged.connect(self.onTextChangedUPH("LP1023BSSM"))
        self.AW1822DR3W.textChanged.connect(self.onTextChangedUPH("AW1822DR3W"))
        self.AS_10TR4RYDTD02.textChanged.connect(
            self.onTextChangedUPH("AS-10TR4RYDTD02")
        )
        self.AS_10TR4RYD02D.textChanged.connect(
            self.onTextChangedUPH("AS-10TR4RYD02(D)")
        )

        self.AS_10TR4RYDTU02.textChanged.connect(
            self.onTextChangedUPH("AS-10TR4RYDTU02")
        )
        self.AS_12CR4RVEDJ01.textChanged.connect(
            self.onTextChangedUPH("AS-12CR4RVEDJ01")
        )
        self.AS_12TR4RYDTU00B.textChanged.connect(
            self.onTextChangedUPH("AS-12TR4RYDTU00B")
        )
        self.AS_18TW4RGATU00.textChanged.connect(
            self.onTextChangedUPH("AS-18TW4RGATU00")
        )
        self.AS_12CR4RVETD01.textChanged.connect(
            self.onTextChangedUPH("AS-12CR4RVETD01")
        )
        self.AS_18TW4RGATD00.textChanged.connect(
            self.onTextChangedUPH("AS-18TW4RGATD00")
        )
        self.AS_12TR4RYDTD00B.textChanged.connect(
            self.onTextChangedUPH("AS-12TR4RYDTD00B")
        )
        self.AS_12TR4RYD00BD.textChanged.connect(
            self.onTextChangedUPH("AS-12TR4RYD00B(D)")
        )
        self.AS_12CR4RVE01D.textChanged.connect(
            self.onTextChangedUPH("AS-12CR4RVE01(D)")
        )
        self.AS_18TW4RGA00D.textChanged.connect(
            self.onTextChangedUPH("AS-18TW4RGA00(D)")
        )
        self.HAP0824TWD.textChanged.connect(self.onTextChangedUPH("HAP0824TWD"))
        self.AS_18TW4RGATD00O.textChanged.connect(
            self.onTextChangedUPH("AS-18TW4RGATD00-O")
        )
        self.AS_12TR4RYDTD00BO.textChanged.connect(
            self.onTextChangedUPH("AS-12TR4RYDTD00B-O")
        )
        self.AS_12CR4RVETD01O.textChanged.connect(
            self.onTextChangedUPH("AS-12CR4RVETD01-O")
        )
        self.AS_10CR4RYDTD02O.textChanged.connect(
            self.onTextChangedUPH("AS-10CR4RYDTD02-O")
        )
        self.AS_10TR4RYDTD02O.textChanged.connect(
            self.onTextChangedUPH("AS-10TR4RYDTD02-O")
        )
        self.RMV08A10A.textChanged.connect(self.onTextChangedUPH("RMV08A10A"))
        self.RMV10A10A.textChanged.connect(self.onTextChangedUPH("RMV10A10A"))
        self.RMV12A10A.textChanged.connect(self.onTextChangedUPH("RMV12A10A"))

        self.prd_start_h1.setDate(
            QDate.fromString(self.data_json["prd_start_h1"], "yyyy-MM-dd")
        )
        self.prd_start_h2.setDate(
            QDate.fromString(self.data_json["prd_start_h2"], "yyyy-MM-dd")
        )
        self.prd_start_h3.setDate(
            QDate.fromString(self.data_json["prd_start_h3"], "yyyy-MM-dd")
        )
        self.prd_start_h4.setDate(
            QDate.fromString(self.data_json["prd_start_h4"], "yyyy-MM-dd")
        )

        # set parameter from json file
        self.man_power_h1.setText(self.data_json["man_power_h1"])
        self.man_power_h2.setText(self.data_json["man_power_h2"])
        self.man_power_h3.setText(self.data_json["man_power_h3"])
        self.man_power_h4.setText(self.data_json["man_power_h4"])
        self.core_processing.setText(self.data_json["core_processing"])
        self.population_line.setText(self.data_json["population_line"])
        self.population_size.setText(self.data_json["population_size"])
        self.generations.setText(self.data_json["generations"])
        self.mutation_rate.setText(self.data_json["mutation_rate"])
        self.file_path.setText(self.data_json["file_path"])
        self.order_path.setText(self.data_json["order_path"])
        self.sale_path.setText(self.data_json["sale_path"])
        self.status.setText(self.data_json["status"])
        self.prd_sheet_name.setText(self.data_json["prd_sheet_name"])
        self.order_sheet_name.setText(self.data_json["order_sheet_name"])
        self.sale_sheet_name.setText(self.data_json["sale_sheet_name"])
        self.change_model.setText(str(self.data_json["change_model"]))
        self.date_diff.setText(str(self.data_json["date_diff"]))
        self.total_ot_day.setText(str(self.data_json["total_ot_day"]))
        self.follow_sale_plan.setText(str(self.data_json["follow_sale_plan"]))
        self.supplier_follow.setText(str(self.data_json["supplier_follow"]))

        self.AP0522CR1W.setText(str(self.data_json["UPH"]["AP0522CR1W"]))
        self.AP0621CR1W.setText(str(self.data_json["UPH"]["AP0621CR1W"]))
        self.AP0722CW1W.setText(str(self.data_json["UPH"]["AP0722CW1W"]))
        self.AP0822CW1W.setText(str(self.data_json["UPH"]["AP0822CW1W"]))
        self.AP1022HW1GD.setText(str(self.data_json["UPH"]["AP1022HW1GD"]))
        self.AP1022CW1G.setText(str(self.data_json["UPH"]["AP1022CW1G"]))
        self.AP1022TW1GD.setText(str(self.data_json["UPH"]["AP1022TW1GD"]))
        self.AP1222CW1W.setText(str(self.data_json["UPH"]["AP1222CW1W"]))
        self.AP1223CW1W.setText(str(self.data_json["UPH"]["AP1223CW1W"]))
        self.AP55023HR1GD.setText(str(self.data_json["UPH"]["AP55023HR1GD"]))
        self.AW1022CW1W.setText(str(self.data_json["UPH"]["AW1022CW1W"]))
        self.AW1023TW1W.setText(str(self.data_json["UPH"]["AW1023TW1W"]))
        self.AW1822CW3W.setText(str(self.data_json["UPH"]["AW1822CW3W"]))
        self.AW2422CW3W.setText(str(self.data_json["UPH"]["AW2422CW3W"]))
        self.AW1822DR3W.setText(str(self.data_json["UPH"]["AW1822DR3W"]))
        self.AW1422CW1W.setText(str(self.data_json["UPH"]["AW1422CW1W"]))
        self.AW1222CW1W.setText(str(self.data_json["UPH"]["AW1222CW1W"]))
        self.AW1222TW1W.setText(str(self.data_json["UPH"]["AW1222TW1W"]))
        self.AW0823CW1W.setText(str(self.data_json["UPH"]["AW0823CW1W"]))
        self.AW1823TW3W.setText(str(self.data_json["UPH"]["AW1823TW3W"]))
        self.AW0822DR1W.setText(str(self.data_json["UPH"]["AW0822DR1W"]))
        self.AW1422TW1W.setText(str(self.data_json["UPH"]["AW1422TW1W"]))
        self.AW2422DR3W.setText(str(self.data_json["UPH"]["AW2422DR3W"]))
        self.AW0823TW1W.setText(str(self.data_json["UPH"]["AW0823TW1W"]))
        self.AW1221DR3W.setText(str(self.data_json["UPH"]["AW1221DR3W"]))
        self.AW2223TW3W.setText(str(self.data_json["UPH"]["AW2223TW3W"]))
        self.LP0823GSSM.setText(str(self.data_json["UPH"]["LP0823GSSM"]))
        self.LP0623WSR.setText(str(self.data_json["UPH"]["LP0623WSR"]))
        self.LP0624WFR.setText(str(self.data_json["UPH"]["LP0624WFR"]))
        self.LP0524WFR.setText(str(self.data_json["UPH"]["LP0524WFR"]))
        self.LP0723WSR.setText(str(self.data_json["UPH"]["LP0723WSR"]))
        self.LP1023BSSM.setText(str(self.data_json["UPH"]["LP1023BSSM"]))
        self.AS_10TR4RYDTD02.setText(str(self.data_json["UPH"]["AS-10TR4RYDTD02"]))
        self.AS_10TR4RYD02D.setText(str(self.data_json["UPH"]["AS-10TR4RYD02(D)"]))
        self.AS_10TR4RYDTU02.setText(str(self.data_json["UPH"]["AS-10TR4RYDTU02"]))
        self.AS_12CR4RVEDJ01.setText(str(self.data_json["UPH"]["AS-12CR4RVEDJ01"]))
        self.AS_12TR4RYDTU00B.setText(str(self.data_json["UPH"]["AS-12TR4RYDTU00B"]))
        self.AS_18TW4RGATU00.setText(str(self.data_json["UPH"]["AS-18TW4RGATU00"]))
        self.AS_12CR4RVETD01.setText(str(self.data_json["UPH"]["AS-12CR4RVETD01"]))
        self.AS_18TW4RGATD00.setText(str(self.data_json["UPH"]["AS-18TW4RGATD00"]))
        self.AS_12TR4RYDTD00B.setText(str(self.data_json["UPH"]["AS-12TR4RYDTD00B"]))
        self.AS_12TR4RYD00BD.setText(str(self.data_json["UPH"]["AS-12TR4RYD00B(D)"]))
        self.AS_12CR4RVE01D.setText(str(self.data_json["UPH"]["AS-12CR4RVE01(D)"]))
        self.AS_18TW4RGA00D.setText(str(self.data_json["UPH"]["AS-18TW4RGA00(D)"]))
        self.HAP0824TWD.setText(str(self.data_json["UPH"]["HAP0824TWD"]))
        self.AS_18TW4RGATD00O.setText(str(self.data_json["UPH"]["AS-18TW4RGATD00-O"]))
        self.AS_12TR4RYDTD00BO.setText(str(self.data_json["UPH"]["AS-12TR4RYDTD00B-O"]))
        self.AS_12CR4RVETD01O.setText(str(self.data_json["UPH"]["AS-12CR4RVETD01-O"]))
        self.AS_10CR4RYDTD02O.setText(str(self.data_json["UPH"]["AS-10CR4RYDTD02-O"]))
        self.AS_10TR4RYDTD02O.setText(str(self.data_json["UPH"]["AS-10TR4RYDTD02-O"]))
        self.RMV08A10A.setText(str(self.data_json["UPH"]["RMV08A10A"]))
        self.RMV10A10A.setText(str(self.data_json["UPH"]["RMV10A10A"]))
        self.RMV12A10A.setText(str(self.data_json["UPH"]["RMV12A10A"]))

        # Brown File
        self.pushButton.clicked.connect(self.getPlanExelFile)
        self.pushButton_2.clicked.connect(self.getOrderExelFile)
        self.pushButton_3.clicked.connect(self.getSaleExelFile)

        # Run Button
        self.pushButton_4.clicked.connect(self.start_heavy_process)
        self.pushButton_5.clicked.connect(self.ClearDataPlanning)

    def setupData(self):
        mapBatch_h1 = {
            "Y1G": self.Y1G_h1,
            "Y1K": self.Y1K_h1,
            "Y1L": self.Y1L_h1,
            "Y1N": self.Y1N_h1,
            "Y1P": self.Y1P_h1,
            "C1A": self.C1A_h1,
            "C1B": self.C1B_h1,
            "C1F": self.C1F_h1,
            "C1U": self.C1U_h1,
            "C1V": self.C1V_h1,
            "C1W": self.C1W_h1,
            "C1Y": self.C1Y_h1,
            "C2X": self.C2X_h1,
            "C3D": self.C3D_h1,
            "LG-Y1L": self.LGY1L_h1,
            "LG-Y1N": self.LGY1N_h1,
            "LG-Y1K": self.LGY1K_h1,
            "W1M": self.W1M_h1,
            "W1R": self.W1R_h1,
            "W1V": self.W1V_h1,
        }
        mapBatch_h2 = {
            "Y1G": self.Y1G_h2,
            "Y1K": self.Y1K_h2,
            "Y1L": self.Y1L_h2,
            "Y1N": self.Y1N_h2,
            "Y1P": self.Y1P_h2,
            "C1A": self.C1A_h2,
            "C1B": self.C1B_h2,
            "C1F": self.C1F_h2,
            "C1U": self.C1U_h2,
            "C1V": self.C1V_h2,
            "C1W": self.C1W_h2,
            "C1Y": self.C1Y_h2,
            "C2X": self.C2X_h2,
            "C3D": self.C3D_h2,
            "LG-Y1L": self.LGY1L_h2,
            "LG-Y1N": self.LGY1N_h2,
            "LG-Y1K": self.LGY1K_h2,
            "W1M": self.W1M_h2,
            "W1R": self.W1R_h2,
            "W1V": self.W1V_h2,
        }
        mapBatch_h3 = {
            "Y1G": self.Y1G_h3,
            "Y1K": self.Y1K_h3,
            "Y1L": self.Y1L_h3,
            "Y1N": self.Y1N_h3,
            "Y1P": self.Y1P_h3,
            "C1A": self.C1A_h3,
            "C1B": self.C1B_h3,
            "C1F": self.C1F_h3,
            "C1U": self.C1U_h3,
            "C1V": self.C1V_h3,
            "C1W": self.C1W_h3,
            "C1Y": self.C1Y_h3,
            "C2X": self.C2X_h3,
            "C3D": self.C3D_h3,
            "LG-Y1L": self.LGY1L_h3,
            "LG-Y1N": self.LGY1N_h3,
            "LG-Y1K": self.LGY1K_h3,
            "W1M": self.W1M_h3,
            "W1R": self.W1R_h3,
            "W1V": self.W1V_h3,
        }
        mapBatch_h4 = {
            "Y1G": self.Y1G_h4,
            "Y1K": self.Y1K_h4,
            "Y1L": self.Y1L_h4,
            "Y1N": self.Y1N_h4,
            "Y1P": self.Y1P_h4,
            "C1A": self.C1A_h4,
            "C1B": self.C1B_h4,
            "C1F": self.C1F_h4,
            "C1U": self.C1U_h4,
            "C1V": self.C1V_h4,
            "C1W": self.C1W_h4,
            "C1Y": self.C1Y_h4,
            "C2X": self.C2X_h4,
            "C3D": self.C3D_h4,
            "LG-Y1N": self.LGY1N_h4,
            "LG-Y1K": self.LGY1K_h4,
            "W1M": self.W1M_h4,
            "W1R": self.W1R_h4,
            "W1V": self.W1V_h4,
        }
        map_line = {
            "HISENSE 1": self.h1,
            "HISENSE 2": self.h2,
            "HISENSE 3": self.h3,
            "HISENSE 4": self.h4,
        }
        map_date = {
            "Monday": self.Monday,
            "Tuesday": self.Tuesday,
            "Wednesday": self.Wednesday,
            "Thursday": self.Thursday,
            "Friday": self.Friday,
            "Saturday": self.Saturday,
        }
        for batch in self.data_json["checkbox_h1"]:
            mapBatch_h1[batch].setChecked(True)
        for batch in self.data_json["checkbox_h2"]:
            mapBatch_h2[batch].setChecked(True)
        for batch in self.data_json["checkbox_h3"]:
            mapBatch_h3[batch].setChecked(True)
        for batch in self.data_json["checkbox_h4"]:
            mapBatch_h4[batch].setChecked(True)
        for line in self.data_json["prd_line"]:
            map_line[line].setChecked(True)
        for day in self.data_json["ot_plan"]:
            map_date[day].setChecked(True)

        if self.data_json["status_plan"] == "CONTINUE PLAN":
            self.continue_plan.setChecked(True)
            self.new_season.setChecked(False)
        elif self.data_json["status_plan"] == "NEW SEASON":
            self.continue_plan.setChecked(False)
            self.new_season.setChecked(True)

    def on_Home_toggled(self):
        self.stackedWidget.setCurrentIndex(0)

    def on_Config_toggled(self):
        self.stackedWidget.setCurrentIndex(1)

    def on_uph_toggled(self):
        self.stackedWidget.setCurrentIndex(2)

    def on_click_prd_line(self, checkbox):
        if checkbox.isChecked():
            text = checkbox.text()
            self.prd_line.append(text)
        else:
            text = checkbox.text()
            self.prd_line.remove(text)
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["prd_line"] = self.prd_line
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        print(self.prd_line)
        return text

    def on_click_checkbox_h1(self, checkbox):
        if checkbox.isChecked():
            text = checkbox.text()
            self.checkbox_h1.append(text)
        else:
            text = checkbox.text()
            self.checkbox_h1.remove(text)
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["checkbox_h1"] = self.checkbox_h1
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        return text

    def on_click_checkbox_h2(self, checkbox):
        if checkbox.isChecked():
            text = checkbox.text()
            self.checkbox_h2.append(text)
        else:
            text = checkbox.text()
            self.checkbox_h2.remove(text)
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["checkbox_h2"] = self.checkbox_h2
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        return text

    def on_click_checkbox_h3(self, checkbox):
        if checkbox.isChecked():
            text = checkbox.text()
            self.checkbox_h3.append(text)
        else:
            text = checkbox.text()
            self.checkbox_h3.remove(text)
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["checkbox_h3"] = self.checkbox_h3
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        return text

    def on_click_checkbox_h4(self, checkbox):
        if checkbox.isChecked():
            text = checkbox.text()
            self.checkbox_h4.append(text)
        else:
            text = checkbox.text()
            self.checkbox_h4.remove(text)
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["checkbox_h4"] = self.checkbox_h4
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        return text

    def ot_day(self, checkbox):
        if checkbox.isChecked():
            text = checkbox.text()
            self.ot_plan.append(text)
        else:
            text = checkbox.text()
            self.ot_plan.remove(text)
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["ot_plan"] = self.ot_plan
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        return text

    def status_plan(self, checkbox):
        if checkbox.isChecked():
            text = checkbox.text()
            if text == "CONTINUE PLAN":
                self.continue_plan.setChecked(True)
                self.new_season.setChecked(False)
            elif text == "NEW SEASON":
                self.continue_plan.setChecked(False)
                self.new_season.setChecked(True)
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["status_plan"] = text
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        return text

    def ClearDataPlanning(self):
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["status"] = "Clear Data"
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        self.status.setText("Clear Data")
        file_path = self.data_json["file_path"]
        sheet_name = "Production plan"
        clear_data = ClearDataPlanning(
            file_path, sheet_name, start_row=11, start_column=18
        )
        clear_data.clear_data_in_range(None, None)  # ล้างข้อมูลตั้งแต่ R11 เป็นต้นไป
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["status"] = "Ready"
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        self.status.setText("Ready")

    def onTextChanged(self, key):
        def inner(text):
            if text:
                try:
                    with open("parameter.json", "r") as json_file:
                        data = json.load(json_file)
                except FileNotFoundError:
                    data = {}
                data[key] = text
                with open("parameter.json", "w") as json_file:
                    json.dump(data, json_file)

        return inner

    def onTextChangedUPH(self, key):
        def inner(text):
            if text:
                try:
                    with open("parameter.json", "r") as json_file:
                        data = json.load(json_file)
                except FileNotFoundError:
                    data = {}
                UPH = data["UPH"]
                UPH[key] = text
                with open("parameter.json", "w") as json_file:
                    json.dump(data, json_file)

        return inner

    def convert_thai_date_to_iso(self, thai_date):
        # แปลงเลขไทยเป็นเลขอารบิก
        thai_digits = ["๐", "๑", "๒", "๓", "๔", "๕", "๖", "๗", "๘", "๙"]
        arabic_digits = [str(i) for i in range(10)]
        for i in range(10):
            thai_date = thai_date.replace(thai_digits[i], arabic_digits[i])

        # แปลงเดือนไทยเป็นเดือนอังกฤษ
        thai_months = [
            "มกราคม",
            "กุมภาพันธ์",
            "มีนาคม",
            "เมษายน",
            "พฤษภาคม",
            "มิถุนายน",
            "กรกฎาคม",
            "สิงหาคม",
            "กันยายน",
            "ตุลาคม",
            "พฤศจิกายน",
            "ธันวาคม",
        ]
        english_months = [str(i).zfill(2) for i in range(1, 13)]
        for i in range(12):
            thai_date = thai_date.replace(thai_months[i], english_months[i])

        # แปลงเป็นรูปแบบ "yyyy-MM-dd"
        date_parts = re.split(r"[-/]", thai_date)
        if len(date_parts) == 3:
            year = int(date_parts[0])
            month = int(date_parts[1])
            day = int(date_parts[2])
            return f"{year:04d}-{month:02d}-{day:02d}"
        return None

    def onDateChanged(self, key, date):
        formatted_date = self.convert_thai_date_to_iso(date.toString("yyyy-MM-dd"))
        try:
            with open("parameter.json", "r") as json_file:
                data = json.load(json_file)
        except FileNotFoundError:
            data = {}
        data[key] = formatted_date
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)

    def getPlanExelFile(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            None,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )

        if file_path:
            self.file_path.setText(file_path)
            try:
                with open("parameter.json", "r") as json_file:
                    data = json.load(json_file)
            except FileNotFoundError:
                data = {}

            data["file_path"] = file_path

            with open("parameter.json", "w") as json_file:
                json.dump(data, json_file)

    def getOrderExelFile(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            None,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )

        if file_path:
            self.order_path.setText(file_path)
            try:
                with open("parameter.json", "r") as json_file:
                    data = json.load(json_file)
            except FileNotFoundError:
                data = {}
            data["order_path"] = file_path
            with open("parameter.json", "w") as json_file:
                json.dump(data, json_file)

    def getSaleExelFile(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            None,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx);;All Files (*)",
            options=options,
        )

        if file_path:
            # print("เส้นทางไฟล์ Excel ที่ถูกเลือก:", file_path)
            self.sale_path.setText(file_path)
            try:
                with open("parameter.json", "r") as json_file:
                    data = json.load(json_file)
            except FileNotFoundError:
                data = {}

            data["sale_path"] = file_path

            with open("parameter.json", "w") as json_file:
                json.dump(data, json_file)
    
    def showErrorAlert(self, error_message):
        QMessageBox.warning(self, "Message Warning", error_message)

if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.setupData()
    ui.on_Home_toggled()
    fetch_thread = threading.Thread(target=ui.fetch_data_periodically, args=(3,))
    fetch_thread.daemon = True  # ทำให้เธรดเป็นเธรดแบ็คกราวนด์ (background thread)
    fetch_thread.start()

    MainWindow.show()
    sys.exit(app.exec_())
