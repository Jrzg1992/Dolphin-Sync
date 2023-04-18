#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023/4/18 17:59
# @Author  : Jrzg


import sys
import json
import queue
import zmq
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, \
    QCheckBox, QHBoxLayout, QLineEdit, QLabel, QTabWidget
from PyQt5.QtGui import QIcon
from datetime import datetime
import configparser
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class MessageReceiverThread(QThread):
    message_received = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.context = zmq.Context()
        self.socket = self.context.socket(zmq.SUB)
        self.socket.connect(addres)
        self.socket.setsockopt(zmq.SUBSCRIBE, b"")

    def run(self):
        while True:
            message = self.socket.recv()
            json_data = message.decode('gbk')
            self.message_received.emit(json_data)


class TableUpdaterThread(QThread):
    table_updated = pyqtSignal(list)

    def __init__(self, queue, parent=None):
        super().__init__(parent)
        self.queue = queue
        self.last_data = None

    def run(self):
        while True:
            json_data = self.queue.get()
            data = json.loads(json_data)
            data = [data]
            # print(datetime.now(), data)
            if self.last_data is None or data != self.last_data:
                self.table_updated.emit(data)
                self.last_data = data.copy()


class MainWidget(QWidget):
    def __init__(self):
        super().__init__()
        icon = QIcon("dolphin.png")
        self.setWindowIcon(icon)
        self.setWindowTitle("Dolphin Sync Client V1.01")
        self.last_data = None
        self.layout = QVBoxLayout(self)

        # Create a QTabWidget to hold multiple tabs
        self.tab_widget = QTabWidget(self)
        self.layout.addWidget(self.tab_widget)

        self.topmost_checkbox = QCheckBox("Always on Top")
        self.topmost_checkbox.stateChanged.connect(self.set_topmost)
        self.layout.addWidget(self.topmost_checkbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("Data Address:"))
        self.address_lineedit = QLineEdit("1. Please modify the service address in the configuration file! "
                                          "2. The program did not undergo sufficient stress testing!", self)
        hbox.addWidget(self.address_lineedit)

        self.layout.addLayout(hbox)

        self.resize(1200, 700)

    def update_table(self, data):
        data = data[0]
        keys = list(data.keys())

        # Iterate over the keys and update the existing tab or create a new tab
        for key in keys:
            tab_index = self.get_tab_index_by_name(key)
            if tab_index != -1:
                table = self.tab_widget.widget(tab_index)
                self.show_data(table, data[key])
            else:
                table = QTableWidget()
                self.tab_widget.addTab(table, key)
                self.show_data(table, data[key])

    def get_tab_index_by_name(self, name):
        for i in range(self.tab_widget.count()):
            if self.tab_widget.tabText(i) == name:
                return i
        return -1

    def show_data(self, table, data):
        table.setRowCount(len(data))
        table.setColumnCount(len(data[0]))

        print(datetime.now(), data)

        # Update the table widget with new data
        for row, item in enumerate(data):
            for col, value in enumerate(item):
                value = '' if value is None else str(value)
                table_item = QTableWidgetItem(value)
                table_item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row, col, table_item)

        # Remember the current data for the next update
        self.last_data = data.copy()

    def set_topmost(self, state):
        if state == Qt.Checked:
            self.setWindowFlags(Qt.WindowStaysOnTopHint)
            self.show()
        else:
            self.setWindowFlags(Qt.Widget)
            self.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    queue = queue.Queue()
    main_window = MainWidget()
    # 创建一个 ConfigParser 对象
    config = configparser.ConfigParser()

    # 读取配置文件
    config.read('config.ini')
    # 获取配置项的值
    addres = config.get('sever_', 'svr_addres')

    message_receiver_thread = MessageReceiverThread()
    message_receiver_thread.message_received.connect(queue.put)
    message_receiver_thread.start()

    table_updater_thread = TableUpdaterThread(queue)
    table_updater_thread.table_updated.connect(main_window.update_table)
    table_updater_thread.start()

    main_window.show()

    sys.exit(app.exec_())
