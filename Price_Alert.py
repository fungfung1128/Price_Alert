import sys
import csv
import time
import datetime
import threading
import json
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTableWidget,
                             QTableWidgetItem, QVBoxLayout, QWidget,
                             QPushButton, QLabel, QTabWidget, QComboBox,
                             QToolBar, QLineEdit, QHBoxLayout, QMessageBox,
                             QMenu, QTimeEdit, QCheckBox)
from PyQt6.QtCore import QTimer, Qt, QTime
from PyQt6.QtGui import QAction, QColor
import pandas as pd
from playsound import playsound
import MetaTrader5 as mt5
from zoneinfo import ZoneInfo


class NonScrollableComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)

    def wheelEvent(self, event):
        event.ignore()


class TimeEdit(QTimeEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDisplayFormat("HH:mm:ss")
        self.setTime(QTime(0, 0, 0))
        self.setMinimumTime(QTime(0, 0, 0))
        self.setMaximumTime(QTime(23, 59, 59))
        self.setStyleSheet("""
            QTimeEdit {
                background-color: white;
                border: 1px solid #ccc;
                padding: 2px;
                min-width: 80px;
            }
            QTimeEdit:focus {
                border: 2px solid #1E90FF;
            }
            QTimeEdit::up-button {
                width: 16px;
                height: 12px;
                subcontrol-origin: border;
                subcontrol-position: top right;
            }
            QTimeEdit::down-button {
                width: 16px;
                height: 12px;
                subcontrol-origin: border;
                subcontrol-position: bottom right;
            }
            QTimeEdit::up-arrow {
                width: 10px;
                height: 10px;
            }
            QTimeEdit::down-arrow {
                width: 10px;
                height: 10px;
            }
        """)

    def setText(self, text):
        try:
            time = QTime.fromString(text, "HH:mm:ss")
            if not time.isValid():
                time = QTime.fromString(text, "h:mm:ss")
            if time.isValid():
                self.setTime(time)
        except:
            pass

    def wheelEvent(self, event):
        event.ignore()


class PriceMonitor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("價格監控")
        self.setGeometry(100, 100, 1400, 800)

        self.tz = ZoneInfo("Asia/Shanghai")
        self.auto_fit_enabled = True
        self.csv_lock = threading.Lock()
        self.config_file = "app_config.json"

        if not mt5.initialize():
            print("無法初始化 MT5")
            sys.exit()

        self.default_symbols = [
            "XAUUSD", "XAGUSD", "PLT", "PAD", "COPPER", "IRON", "GAUCNH", "XALUSD",
            "HKGHKD", "XNIUSD", "XZNUSD", "AUDUSD", "EURUSD", "GBPUSD", "NZDUSD",
            "USDCAD", "USDCHF", "USDJPY", "EURJPY", "NZDJPY", "GBPJPY", "CADJPY",
            "AUDJPY", "EURCHF", "EURGBP", "EURAUD", "GBPCHF", "GBPAUD", "AUDNZD",
            "EURCAD", "EURNZD", "GBPCAD", "AUDCAD", "NZDCAD", "USDCNH", "HKDCNH",
            "USOil", "UKOil", "NGAS", "CHINA300", "HK50", "JPN225", "A50", "STI",
            "AS200", "INDIA50", "KS200", "SH50", "VN30", "DJ30", "SP500", "TECH100",
            "RUSS2000", "USDINDEX", "GER30", "FRA40", "UK100", "EUR50", "EUR600",
            "AEX25", "SOYBEAN", "CORN", "WHEAT", "COCOA", "COFFEE", "SUGAR", "COTTON",
            "00005.HK", "AAPL", "BTCUSDT", "ETHUSDT"
        ]

        self.symbols = []
        self.wav_paths = {}
        self.wav_paths_resume = {}
        self.custom_times = {}
        self.alert_enabled = {}
        self.order_options = {}
        self.symbol_digits = {}  # 新增：儲存每個交易品種的小數位數

        self.load_config()
        self.load_symbols_from_file()
        self.load_symbol_digits()  # 新增：載入每個品種的 digits
        self.setup_ui()
        self.load_last_parameters()

        self.last_prices = {}
        self.running = False
        self.csv_file = "MT5_Latest_Price_All.csv"

        self.price_thread = None
        self.refresh_timer = QTimer()
        self.refresh_timer.timeout.connect(self.update_display)
        self.status_label.setText("狀態: 已停止")

    def load_config(self):
        if Path(self.config_file).exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.last_product_path = config.get('product_path', str(Path('product_list.csv').absolute()))
                    self.last_clock = config.get('clock', None)
            except Exception as e:
                print(f"載入配置錯誤: {e}")
                self.last_product_path = str(Path('product_list.csv').absolute())
                self.last_clock = None
        else:
            self.last_product_path = str(Path('product_list.csv').absolute())
            self.last_clock = None

    def save_config(self):
        config = {
            'product_path': self.product_path_input.text().strip(),
            'clock': self.clock_combo.currentText() if self.clock_combo.count() > 0 else None
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"保存配置錯誤: {e}")

    def load_symbols_from_file(self):
        product_file = self.last_product_path if hasattr(self, 'last_product_path') else "product_list.csv"
        if Path(product_file).exists():
            try:
                with open(product_file, 'r') as f:
                    reader = csv.reader(f)
                    self.symbols = []
                    self.wav_paths.clear()
                    self.wav_paths_resume.clear()
                    self.alert_enabled.clear()
                    self.custom_times.clear()
                    for row in reader:
                        if row and row[0].strip():
                            symbol = row[0].strip()
                            self.symbols.append(symbol)
                            self.wav_paths[symbol] = row[1].strip() if len(row) > 1 else ""
                            self.wav_paths_resume[symbol] = row[2].strip() if len(row) > 2 else ""
                            self.alert_enabled[symbol] = True
                            self.custom_times[symbol] = "00:00:00"
            except Exception as e:
                print(f"載入產品列表錯誤: {e}。使用預設符號。")
                self.symbols = self.default_symbols.copy()
                self.wav_paths = {symbol: "" for symbol in self.default_symbols}
                self.wav_paths_resume = {symbol: "" for symbol in self.default_symbols}
                self.alert_enabled = {symbol: True for symbol in self.default_symbols}
                self.custom_times = {symbol: "00:00:00" for symbol in self.default_symbols}
        else:
            print(f"找不到產品檔案 '{product_file}'。使用預設符號。")
            self.symbols = self.default_symbols.copy()
            self.wav_paths = {symbol: "" for symbol in self.default_symbols}
            self.wav_paths_resume = {symbol: "" for symbol in self.default_symbols}
            self.alert_enabled = {symbol: True for symbol in self.default_symbols}
            self.custom_times = {symbol: "00:00:00" for symbol in self.default_symbols}

    def load_symbol_digits(self):
        """從 MT5 獲取每個交易品種的小數位數"""
        for symbol in self.symbols:
            try:
                mt5.symbol_select(symbol, True)
                info = mt5.symbol_info(symbol)
                if info is not None:
                    self.symbol_digits[symbol] = info.digits
                else:
                    self.symbol_digits[symbol] = 5  # 預設 5 位小數
            except Exception as e:
                print(f"無法獲取 {symbol} 的 digits: {e}")
                self.symbol_digits[symbol] = 5  # 預設值

    def load_product_list(self):
        file_path = self.product_path_input.text().strip() or "product_list.csv"
        if not Path(file_path).exists():
            QMessageBox.warning(self, "警告", f"檔案不存在: {file_path}")
            return

        try:
            with open(file_path, 'r') as f:
                reader = csv.reader(f)
                new_symbols = []
                self.wav_paths.clear()
                self.wav_paths_resume.clear()
                self.alert_enabled.clear()
                old_custom_times = self.custom_times.copy()
                self.custom_times.clear()

                for row in reader:
                    if row and row[0].strip():
                        symbol = row[0].strip()
                        new_symbols.append(symbol)
                        self.wav_paths[symbol] = row[1].strip() if len(row) > 1 else ""
                        self.wav_paths_resume[symbol] = row[2].strip() if len(row) > 2 else ""
                        self.alert_enabled[symbol] = True
                        self.custom_times[symbol] = old_custom_times.get(symbol, "00:00:00")

                if not new_symbols:
                    QMessageBox.warning(self, "警告", "檔案為空或無有效符號!")
                    return

                self.symbols = new_symbols
                self.load_symbol_digits()  # 更新 digits
                self.rebuild_price_table()
                self.rebuild_schedule_table()
                self.product_table.clearContents()
                self.product_table.setRowCount(len(self.symbols))
                for i, symbol in enumerate(self.symbols):
                    self.product_table.setItem(i, 0, QTableWidgetItem(symbol))
                    wav_item_stop = QTableWidgetItem(self.wav_paths[symbol])
                    wav_item_stop.setFlags(wav_item_stop.flags() | Qt.ItemFlag.ItemIsEditable)
                    self.product_table.setItem(i, 1, wav_item_stop)
                    wav_item_resume = QTableWidgetItem(self.wav_paths_resume[symbol])
                    wav_item_resume.setFlags(wav_item_resume.flags() | Qt.ItemFlag.ItemIsEditable)
                    self.product_table.setItem(i, 2, wav_item_resume)

                QMessageBox.information(self, "成功", f"已從 '{file_path}' 載入產品列表")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"載入產品列表失敗: {str(e)}")

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        self.create_toolbar()

        control_layout = QHBoxLayout()
        control_layout.addStretch()

        self.start_button = QPushButton("開始")
        self.stop_button = QPushButton("停止")
        self.start_button.setFixedSize(80, 30)
        self.stop_button.setFixedSize(80, 30)
        self.start_button.clicked.connect(self.start_monitoring)
        self.stop_button.clicked.connect(self.stop_monitoring)
        control_layout.addWidget(self.start_button)
        control_layout.addWidget(self.stop_button)

        path_layout = QHBoxLayout()
        path_layout.addStretch()

        product_path_layout = QHBoxLayout()
        product_path_layout.addWidget(QLabel("產品列表:"))
        self.product_path_input = QLineEdit(
            self.last_product_path if hasattr(self, 'last_product_path') else str(Path('product_list.csv').absolute()))
        self.product_path_input.setMinimumWidth(300)
        product_path_layout.addWidget(self.product_path_input)

        params_path_layout = QHBoxLayout()
        params_path_layout.addWidget(QLabel("參數:"))
        self.params_path_input = QLineEdit(f"{Path('trading_schedule_parameters.csv').absolute()}")
        self.params_path_input.setMinimumWidth(300)
        params_path_layout.addWidget(self.params_path_input)

        path_layout.addLayout(product_path_layout)
        path_layout.addSpacing(10)
        path_layout.addLayout(params_path_layout)

        main_layout.addLayout(control_layout)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.create_price_tab()
        self.create_schedule_tab()
        self.create_product_list_tab()

        main_layout.addLayout(path_layout)

    def load_last_parameters(self):
        if hasattr(self, 'last_clock') and self.last_clock and self.last_clock in self.order_options:
            param_path = self.order_options[self.last_clock]
            if Path(param_path).exists():
                self.apply_clock_parameters_from_file(param_path)
                self.clock_combo.setCurrentText(self.last_clock)

    def apply_clock_parameters_from_file(self, param_path):
        try:
            with open(param_path, 'r') as f:
                reader = csv.reader(f)
                headers = next(reader)
                for row in reader:
                    if not row or len(row) < 13:
                        continue
                    symbol = row[0].strip()
                    if symbol not in self.symbols:
                        continue

                    row_idx = self.symbols.index(symbol)

                    time_tolerance = QTime.fromString(row[1].strip(), "HH:mm:ss")
                    if not time_tolerance.isValid():
                        time_tolerance = QTime.fromString(row[1].strip(), "h:mm:ss")
                    if time_tolerance.isValid():
                        self.price_table.cellWidget(row_idx, 6).setTime(time_tolerance)
                        self.custom_times[symbol] = time_tolerance.toString("HH:mm:ss")

                    self.schedule_table.setItem(row_idx, 1, QTableWidgetItem(row[2]))
                    self.schedule_table.cellWidget(row_idx, 2).setCurrentText(row[3])

                    start_time = QTime.fromString(row[4].strip(), "HH:mm:ss")
                    if not start_time.isValid():
                        start_time = QTime.fromString(row[4].strip(), "h:mm:ss")
                    if start_time.isValid():
                        self.schedule_table.cellWidget(row_idx, 3).setTime(start_time)

                    self.schedule_table.cellWidget(row_idx, 4).setCurrentText(row[5])

                    end_time = QTime.fromString(row[6].strip(), "HH:mm:ss")
                    if not end_time.isValid():
                        end_time = QTime.fromString(row[6].strip(), "h:mm:ss")
                    if end_time.isValid():
                        self.schedule_table.cellWidget(row_idx, 5).setTime(end_time)

                    for col, value in zip(range(6, 12), row[7:13]):
                        break_time = QTime.fromString(value.strip(), "HH:mm:ss")
                        if not break_time.isValid():
                            break_time = QTime.fromString(value.strip(), "h:mm:ss")
                        if break_time.isValid():
                            self.schedule_table.cellWidget(row_idx, col).setTime(break_time)
        except Exception as e:
            print(f"載入參數失敗 {param_path}: {str(e)}")

    def create_price_tab(self):
        price_tab = QWidget()
        price_layout = QVBoxLayout(price_tab)

        self.status_label = QLabel("最後更新: 尚未更新 | 狀態: 已停止")
        price_layout.addWidget(self.status_label)

        self.price_table = QTableWidget()
        self.price_table.setColumnCount(10)
        self.price_table.setRowCount(len(self.symbols))
        self.price_table.setHorizontalHeaderLabels(
            ["", "交易品種", "時間", "買價", "賣價", "時間差", "時間容忍度", "狀態", "警報", "語音警報"])

        for i, symbol in enumerate(self.symbols):
            checkbox = QCheckBox()
            checkbox.setChecked(self.alert_enabled.get(symbol, True))
            checkbox.stateChanged.connect(
                lambda state, s=symbol: self.alert_enabled.__setitem__(s, state == Qt.CheckState.Checked.value))
            self.price_table.setCellWidget(i, 0, checkbox)

            self.price_table.setItem(i, 1, QTableWidgetItem(symbol))

            time_tolerance_edit = TimeEdit()
            time_tolerance_edit.setTime(QTime.fromString(self.custom_times.get(symbol, "00:00:00"), "HH:mm:ss"))
            self.price_table.setCellWidget(i, 6, time_tolerance_edit)

            status_item = QTableWidgetItem("關閉")
            status_item.setBackground(QColor(255, 200, 200))
            self.price_table.setItem(i, 7, status_item)

            alert_item = QTableWidgetItem("關閉")
            alert_item.setBackground(QColor(255, 200, 200))
            self.price_table.setItem(i, 8, alert_item)

            voice_alert_item = QTableWidgetItem("")
            voice_alert_item.setFlags(voice_alert_item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.price_table.setItem(i, 9, voice_alert_item)

        self.price_table.setColumnWidth(9, 100)
        price_layout.addWidget(self.price_table)
        self.tabs.addTab(price_tab, "價格監控")

    def create_schedule_tab(self):
        schedule_tab = QWidget()
        schedule_layout = QVBoxLayout(schedule_tab)

        self.schedule_table = QTableWidget()
        self.schedule_table.setColumnCount(12)
        self.schedule_table.setRowCount(len(self.symbols))
        self.schedule_table.setHorizontalHeaderLabels([
            "交易品種", "產品組", "開始日", "開始時間", "結束日", "結束時間",
            "休息1開始", "休息1結束", "休息2開始", "休息2結束",
            "休息3開始", "休息3結束"
        ])

        days = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

        for i, symbol in enumerate(self.symbols):
            self.schedule_table.setItem(i, 0, QTableWidgetItem(symbol))

            item = QTableWidgetItem("")
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.schedule_table.setItem(i, 1, item)

            start_day_combo = NonScrollableComboBox()
            start_day_combo.addItems(days)
            self.schedule_table.setCellWidget(i, 2, start_day_combo)

            start_time_edit = TimeEdit()
            start_time_edit.setTime(QTime(0, 0, 0))
            self.schedule_table.setCellWidget(i, 3, start_time_edit)

            end_day_combo = NonScrollableComboBox()
            end_day_combo.addItems(days)
            self.schedule_table.setCellWidget(i, 4, end_day_combo)

            end_time_edit = TimeEdit()
            end_time_edit.setTime(QTime(0, 0, 0))
            self.schedule_table.setCellWidget(i, 5, end_time_edit)

            for col in range(6, 12, 2):
                start_edit = TimeEdit()
                start_edit.setTime(QTime(0, 0, 0))
                self.schedule_table.setCellWidget(i, col, start_edit)

                end_edit = TimeEdit()
                end_edit.setTime(QTime(0, 0, 0))
                self.schedule_table.setCellWidget(i, col + 1, end_edit)

        schedule_layout.addWidget(self.schedule_table)

        clock_layout = QHBoxLayout()
        clock_layout.addStretch()

        clock_label = QLabel("轉令:")
        clock_layout.addWidget(clock_label)

        self.clock_combo = NonScrollableComboBox()
        self.load_clock_options()
        clock_layout.addWidget(self.clock_combo)

        self.clock_apply_button = QPushButton("應用")
        self.clock_apply_button.clicked.connect(self.apply_clock_parameters)
        clock_layout.addWidget(self.clock_apply_button)

        clock_layout.addStretch()
        schedule_layout.addLayout(clock_layout)

        self.tabs.addTab(schedule_tab, "交易時間表")

    def create_product_list_tab(self):
        product_tab = QWidget()
        product_layout = QVBoxLayout(product_tab)

        self.product_table = QTableWidget()
        self.product_table.setColumnCount(3)
        self.product_table.setRowCount(len(self.symbols))
        self.product_table.setHorizontalHeaderLabels(["交易品種", "WAV路徑（停價）", "WAV路徑（報價恢復）"])
        self.product_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.product_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)

        for i, symbol in enumerate(self.symbols):
            self.product_table.setItem(i, 0, QTableWidgetItem(symbol))
            wav_item_stop = QTableWidgetItem(self.wav_paths.get(symbol, ""))
            wav_item_stop.setFlags(wav_item_stop.flags() | Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(i, 1, wav_item_stop)
            wav_item_resume = QTableWidgetItem(self.wav_paths_resume.get(symbol, ""))
            wav_item_resume.setFlags(wav_item_resume.flags() | Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(i, 2, wav_item_resume)

        product_layout.addWidget(self.product_table)

        button_layout = QHBoxLayout()

        # 新增 "載入產品" 按鈕到左側
        self.load_products_button = QPushButton("載入產品")
        self.load_products_button.clicked.connect(self.load_product_list)
        button_layout.addWidget(self.load_products_button)

        button_layout.addStretch()  # 將其他按鈕推到右側

        self.add_button = QPushButton("新增產品")
        self.remove_button = QPushButton("移除選中")
        self.move_up_button = QPushButton("上移")
        self.move_down_button = QPushButton("下移")
        self.apply_button = QPushButton("應用變更")

        self.add_button.clicked.connect(self.add_product)
        self.remove_button.clicked.connect(self.remove_product)
        self.move_up_button.clicked.connect(self.move_product_up)
        self.move_down_button.clicked.connect(self.move_product_down)
        self.apply_button.clicked.connect(self.apply_product_changes)

        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addWidget(self.move_up_button)
        button_layout.addWidget(self.move_down_button)
        button_layout.addWidget(self.apply_button)

        product_layout.addLayout(button_layout)
        self.tabs.addTab(product_tab, "產品列表")

    def rebuild_price_table(self):
        old_tolerance = {symbol: self.custom_times.get(symbol, "00:00:00") for symbol in self.symbols}
        self.price_table.clearContents()
        self.price_table.setRowCount(len(self.symbols))

        for i, symbol in enumerate(self.symbols):
            checkbox = QCheckBox()
            checkbox.setChecked(self.alert_enabled.get(symbol, True))
            checkbox.stateChanged.connect(
                lambda state, s=symbol: self.alert_enabled.__setitem__(s, state == Qt.CheckState.Checked.value))
            self.price_table.setCellWidget(i, 0, checkbox)

            self.price_table.setItem(i, 1, QTableWidgetItem(symbol))

            time_tolerance_edit = TimeEdit()
            time_tolerance_edit.setTime(QTime.fromString(old_tolerance.get(symbol, "00:00:00"), "HH:mm:ss"))
            self.price_table.setCellWidget(i, 6, time_tolerance_edit)

            status_item = QTableWidgetItem("關閉")
            status_item.setBackground(QColor(255, 200, 200))
            self.price_table.setItem(i, 7, status_item)

            alert_item = QTableWidgetItem("關閉")
            alert_item.setBackground(QColor(255, 200, 200))
            self.price_table.setItem(i, 8, alert_item)

            voice_alert_item = QTableWidgetItem("")
            voice_alert_item.setFlags(voice_alert_item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.price_table.setItem(i, 9, voice_alert_item)

        self.price_table.setColumnWidth(9, 100)
        self.price_table.resizeColumnsToContents()
        self.price_table.resizeRowsToContents()

    def rebuild_schedule_table(self):
        old_schedule = {}
        for i in range(self.schedule_table.rowCount()):
            symbol = self.schedule_table.item(i, 0).text()
            product_group = self.schedule_table.item(i, 1).text() if self.schedule_table.item(i, 1) else ""
            start_day = self.schedule_table.cellWidget(i, 2).currentText()
            start_time = self.schedule_table.cellWidget(i, 3).time().toString("HH:mm:ss")
            end_day = self.schedule_table.cellWidget(i, 4).currentText()
            end_time = self.schedule_table.cellWidget(i, 5).time().toString("HH:mm:ss")
            breaks = [(self.schedule_table.cellWidget(i, col).time().toString("HH:mm:ss"),
                       self.schedule_table.cellWidget(i, col + 1).time().toString("HH:mm:ss"))
                      for col in range(6, 12, 2)]
            old_schedule[symbol] = (product_group, start_day, start_time, end_day, end_time, breaks)

        self.schedule_table.clearContents()
        self.schedule_table.setRowCount(len(self.symbols))

        days = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]

        for i, symbol in enumerate(self.symbols):
            self.schedule_table.setItem(i, 0, QTableWidgetItem(symbol))

            item = QTableWidgetItem("")
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.schedule_table.setItem(i, 1, item)

            start_day_combo = NonScrollableComboBox()
            start_day_combo.addItems(days)
            self.schedule_table.setCellWidget(i, 2, start_day_combo)

            start_time_edit = TimeEdit()
            start_time_edit.setTime(QTime(0, 0, 0))
            self.schedule_table.setCellWidget(i, 3, start_time_edit)

            end_day_combo = NonScrollableComboBox()
            end_day_combo.addItems(days)
            self.schedule_table.setCellWidget(i, 4, end_day_combo)

            end_time_edit = TimeEdit()
            end_time_edit.setTime(QTime(0, 0, 0))
            self.schedule_table.setCellWidget(i, 5, end_time_edit)

            for col in range(6, 12, 2):
                start_edit = TimeEdit()
                start_edit.setTime(QTime(0, 0, 0))
                self.schedule_table.setCellWidget(i, col, start_edit)

                end_edit = TimeEdit()
                end_edit.setTime(QTime(0, 0, 0))
                self.schedule_table.setCellWidget(i, col + 1, end_edit)

            if symbol in old_schedule:
                product_group, start_day, start_time, end_day, end_time, breaks = old_schedule[symbol]
                self.schedule_table.item(i, 1).setText(product_group)
                self.schedule_table.cellWidget(i, 2).setCurrentText(start_day)
                self.schedule_table.cellWidget(i, 3).setTime(QTime.fromString(start_time, "HH:mm:ss"))
                self.schedule_table.cellWidget(i, 4).setCurrentText(end_day)
                self.schedule_table.cellWidget(i, 5).setTime(QTime.fromString(end_time, "HH:mm:ss"))
                for col, (break_start, break_end) in zip(range(6, 12, 2), breaks):
                    self.schedule_table.cellWidget(i, col).setTime(QTime.fromString(break_start, "HH:mm:ss"))
                    self.schedule_table.cellWidget(i, col + 1).setTime(QTime.fromString(break_end, "HH:mm:ss"))

        self.schedule_table.resizeColumnsToContents()
        self.schedule_table.resizeRowsToContents()

    def save_product_list(self):
        file_path = self.product_path_input.text().strip() or "product_list.csv"
        try:
            with open(file_path, 'w', newline='') as f:
                writer = csv.writer(f)
                for row in range(self.product_table.rowCount()):
                    symbol = self.product_table.item(row, 0).text()
                    wav_path_stop = self.product_table.item(row, 1).text() if self.product_table.item(row, 1) else ""
                    wav_path_resume = self.product_table.item(row, 2).text() if self.product_table.item(row, 2) else ""
                    if symbol:
                        writer.writerow([symbol, wav_path_stop, wav_path_resume])
            QMessageBox.information(self, "成功", f"產品列表已保存至 '{file_path}'")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"保存產品列表失敗: {str(e)}")

    def save_parameters(self):
        data = []
        for row in range(self.schedule_table.rowCount()):
            row_data = []
            row_data.append(self.schedule_table.item(row, 0).text())
            time_tolerance_edit = self.price_table.cellWidget(row, 6)
            row_data.append(time_tolerance_edit.time().toString("HH:mm:ss") if time_tolerance_edit else "00:00:00")
            row_data.append(self.schedule_table.item(row, 1).text() if self.schedule_table.item(row, 1) else "")
            row_data.append(self.schedule_table.cellWidget(row, 2).currentText())
            start_time_edit = self.schedule_table.cellWidget(row, 3)
            row_data.append(start_time_edit.time().toString("HH:mm:ss") if start_time_edit else "00:00:00")
            row_data.append(self.schedule_table.cellWidget(row, 4).currentText())
            end_time_edit = self.schedule_table.cellWidget(row, 5)
            row_data.append(end_time_edit.time().toString("HH:mm:ss") if end_time_edit else "00:00:00")
            for col in range(6, 12):
                time_edit = self.schedule_table.cellWidget(row, col)
                row_data.append(time_edit.time().toString("HH:mm:ss") if time_edit else "00:00:00")
            data.append(row_data)

        selected_clock = self.clock_combo.currentText()
        if selected_clock in self.order_options:
            file_path = self.order_options[selected_clock]
        else:
            file_path = self.params_path_input.text().strip() or "trading_schedule_parameters.csv"

        Path(file_path).parent.mkdir(parents=True, exist_ok=True)

        try:
            with open(file_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["交易品種", "時間容忍度", "產品組", "開始日", "開始時間",
                                 "結束日", "結束時間", "休息1開始", "休息1結束",
                                 "休息2開始", "休息2結束", "休息3開始", "休息3結束"])
                writer.writerows(data)
            QMessageBox.information(self, "成功", f"參數已保存至 '{file_path}'")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"保存參數失敗: {str(e)}")

    def add_product(self):
        row_count = self.product_table.rowCount()
        self.product_table.insertRow(row_count)
        self.product_table.setItem(row_count, 0, QTableWidgetItem("新品種"))
        wav_item_stop = QTableWidgetItem("")
        wav_item_stop.setFlags(wav_item_stop.flags() | Qt.ItemFlag.ItemIsEditable)
        self.product_table.setItem(row_count, 1, wav_item_stop)
        wav_item_resume = QTableWidgetItem("")
        wav_item_resume.setFlags(wav_item_resume.flags() | Qt.ItemFlag.ItemIsEditable)
        self.product_table.setItem(row_count, 2, wav_item_resume)

    def remove_product(self):
        current_row = self.product_table.currentRow()
        if current_row >= 0:
            self.product_table.removeRow(current_row)

    def move_product_up(self):
        current_row = self.product_table.currentRow()
        if current_row > 0:
            symbol = self.product_table.item(current_row, 0).text()
            wav_path_stop = self.product_table.item(current_row, 1).text()
            wav_path_resume = self.product_table.item(current_row, 2).text()
            self.product_table.removeRow(current_row)
            self.product_table.insertRow(current_row - 1)
            self.product_table.setItem(current_row - 1, 0, QTableWidgetItem(symbol))
            wav_item_stop = QTableWidgetItem(wav_path_stop)
            wav_item_stop.setFlags(wav_item_stop.flags() | Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(current_row - 1, 1, wav_item_stop)
            wav_item_resume = QTableWidgetItem(wav_path_resume)
            wav_item_resume.setFlags(wav_item_resume.flags() | Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(current_row - 1, 2, wav_item_resume)
            self.product_table.setCurrentCell(current_row - 1, 0)

    def move_product_down(self):
        current_row = self.product_table.currentRow()
        if current_row >= 0 and current_row < self.product_table.rowCount() - 1:
            symbol = self.product_table.item(current_row, 0).text()
            wav_path_stop = self.product_table.item(current_row, 1).text()
            wav_path_resume = self.product_table.item(current_row, 2).text()
            self.product_table.removeRow(current_row)
            self.product_table.insertRow(current_row + 1)
            self.product_table.setItem(current_row + 1, 0, QTableWidgetItem(symbol))
            wav_item_stop = QTableWidgetItem(wav_path_stop)
            wav_item_stop.setFlags(wav_item_stop.flags() | Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(current_row + 1, 1, wav_item_stop)
            wav_item_resume = QTableWidgetItem(wav_path_resume)
            wav_item_resume.setFlags(wav_item_resume.flags() | Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(current_row + 1, 2, wav_item_resume)
            self.product_table.setCurrentCell(current_row + 1, 0)

    def apply_product_changes(self):
        new_symbols = []
        self.wav_paths.clear()
        self.wav_paths_resume.clear()
        self.alert_enabled.clear()
        old_custom_times = self.custom_times.copy()
        self.custom_times.clear()

        for row in range(self.product_table.rowCount()):
            symbol_item = self.product_table.item(row, 0)
            wav_item_stop = self.product_table.item(row, 1)
            wav_item_resume = self.product_table.item(row, 2)
            if symbol_item and symbol_item.text():
                symbol = symbol_item.text()
                new_symbols.append(symbol)
                self.wav_paths[symbol] = wav_item_stop.text() if wav_item_stop else ""
                self.wav_paths_resume[symbol] = wav_item_resume.text() if wav_item_resume else ""
                self.alert_enabled[symbol] = True
                self.custom_times[symbol] = old_custom_times.get(symbol, "00:00:00")

        if not new_symbols:
            QMessageBox.warning(self, "警告", "產品列表不能為空!")
            return

        self.symbols = new_symbols
        self.load_symbol_digits()  # 更新 digits
        self.save_product_list()
        self.rebuild_price_table()
        self.rebuild_schedule_table()

        QMessageBox.information(self, "成功", "產品列表更新成功!")

    def load_clock_options(self):
        clock_file = "clock_change.txt"
        self.order_options = {}

        try:
            if Path(clock_file).exists():
                with open(clock_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    self.clock_combo.clear()
                    for line in lines:
                        line = line.strip()
                        if line and ',' in line:
                            clock_name, param_path = line.split(',', 1)
                            self.order_options[clock_name.strip()] = param_path.strip()
                            self.clock_combo.addItem(clock_name.strip())
            if self.clock_combo.count() == 0:
                self.clock_combo.addItem("無可用轉令")
        except Exception as e:
            print(f"載入轉令選項失敗: {str(e)}")
            self.clock_combo.addItem("載入失敗")

    def apply_clock_parameters(self):
        selected_clock = self.clock_combo.currentText()
        if selected_clock in self.order_options:
            param_path = self.order_options[selected_clock]
            if Path(param_path).exists():
                try:
                    old_params = {}
                    for i in range(self.schedule_table.rowCount()):
                        symbol = self.schedule_table.item(i, 0).text()
                        time_tolerance = self.price_table.cellWidget(i, 6).time().toString("HH:mm:ss")
                        product_group = self.schedule_table.item(i, 1).text() if self.schedule_table.item(i, 1) else ""
                        start_day = self.schedule_table.cellWidget(i, 2).currentText()
                        start_time = self.schedule_table.cellWidget(i, 3).time().toString("HH:mm:ss")
                        end_day = self.schedule_table.cellWidget(i, 4).currentText()
                        end_time = self.schedule_table.cellWidget(i, 5).time().toString("HH:mm:ss")
                        breaks = [(self.schedule_table.cellWidget(i, col).time().toString("HH:mm:ss"),
                                   self.schedule_table.cellWidget(i, col + 1).time().toString("HH:mm:ss"))
                                  for col in range(6, 12, 2)]
                        old_params[symbol] = (time_tolerance, product_group, start_day, start_time,
                                              end_day, end_time, breaks)

                    with open(param_path, 'r') as f:
                        reader = csv.reader(f)
                        headers = next(reader)
                        for row in reader:
                            if not row or len(row) < 13:
                                continue
                            symbol = row[0].strip()
                            if symbol not in self.symbols:
                                continue

                            row_idx = self.symbols.index(symbol)

                            time_tolerance = QTime.fromString(row[1].strip(), "HH:mm:ss")
                            if not time_tolerance.isValid():
                                time_tolerance = QTime.fromString(row[1].strip(), "h:mm:ss")
                            if time_tolerance.isValid():
                                self.price_table.cellWidget(row_idx, 6).setTime(time_tolerance)
                                self.custom_times[symbol] = time_tolerance.toString("HH:mm:ss")

                            self.schedule_table.setItem(row_idx, 1, QTableWidgetItem(row[2]))
                            self.schedule_table.cellWidget(row_idx, 2).setCurrentText(row[3])

                            start_time = QTime.fromString(row[4].strip(), "HH:mm:ss")
                            if not start_time.isValid():
                                start_time = QTime.fromString(row[4].strip(), "h:mm:ss")
                            if start_time.isValid():
                                self.schedule_table.cellWidget(row_idx, 3).setTime(start_time)

                            self.schedule_table.cellWidget(row_idx, 4).setCurrentText(row[5])

                            end_time = QTime.fromString(row[6].strip(), "HH:mm:ss")
                            if not end_time.isValid():
                                end_time = QTime.fromString(row[6].strip(), "h:mm:ss")
                            if end_time.isValid():
                                self.schedule_table.cellWidget(row_idx, 5).setTime(end_time)

                            for col, value in zip(range(6, 12), row[7:13]):
                                break_time = QTime.fromString(value.strip(), "HH:mm:ss")
                                if not break_time.isValid():
                                    break_time = QTime.fromString(value.strip(), "h:mm:ss")
                                if break_time.isValid():
                                    self.schedule_table.cellWidget(row_idx, col).setTime(break_time)

                    QMessageBox.information(self, "成功", f"已應用轉令 '{selected_clock}' 的參數")
                except Exception as e:
                    QMessageBox.critical(self, "錯誤", f"應用轉令參數失敗: {str(e)}")
            else:
                QMessageBox.warning(self, "警告", f"參數檔案不存在: {param_path}")
        else:
            QMessageBox.warning(self, "警告", "請選擇有效的轉令")

    def update_display(self):
        try:
            with self.csv_lock:
                if Path(self.csv_file).exists():
                    df = pd.read_csv(self.csv_file)
                    current_time = datetime.datetime.now(tz=self.tz)

                    for i, row in df.iterrows():
                        row_idx = self.symbols.index(row['Symbol'])
                        symbol_time = datetime.datetime.strptime(row['Time'], '%Y-%m-%d %H:%M:%S').replace(
                            tzinfo=self.tz)

                        self.price_table.setItem(row_idx, 2, QTableWidgetItem(row['Time']))
                        # 格式化 Bid 和 Ask 根據品種的 digits
                        digits = self.symbol_digits.get(row['Symbol'], 5)
                        self.price_table.setItem(row_idx, 3, QTableWidgetItem(f"{row['Bid']:.{digits}f}"))
                        self.price_table.setItem(row_idx, 4, QTableWidgetItem(f"{row['Ask']:.{digits}f}"))

                        time_diff = current_time - symbol_time
                        if time_diff.total_seconds() < 0:
                            time_diff = datetime.timedelta(seconds=0)
                        time_diff_str = str(time_diff).split('.')[0]
                        self.price_table.setItem(row_idx, 5, QTableWidgetItem(time_diff_str))

                        is_trading_now = self.is_trading(row_idx, current_time)
                        status_item = self.price_table.item(row_idx, 7)
                        if not status_item:
                            status_item = QTableWidgetItem()
                            self.price_table.setItem(row_idx, 7, status_item)

                        if is_trading_now:
                            status_item.setText("交易中")
                            status_item.setBackground(Qt.GlobalColor.green)
                        else:
                            status_item.setText("關閉")
                            status_item.setBackground(QColor(255, 200, 200))

                        alert_item = self.price_table.item(row_idx, 8)
                        if not alert_item:
                            alert_item = QTableWidgetItem()
                            self.price_table.setItem(row_idx, 8, alert_item)

                        time_tolerance_edit = self.price_table.cellWidget(row_idx, 6)
                        tolerance_time = QTime.fromString(time_tolerance_edit.time().toString("HH:mm:ss"), "HH:mm:ss")
                        time_diff_seconds = time_diff.total_seconds()
                        tolerance_seconds = tolerance_time.hour() * 3600 + tolerance_time.minute() * 60 + tolerance_time.second()

                        symbol = row['Symbol']
                        voice_alert_item = self.price_table.item(row_idx, 9)
                        if not voice_alert_item:
                            voice_alert_item = QTableWidgetItem("")
                            voice_alert_item.setFlags(voice_alert_item.flags() | Qt.ItemFlag.ItemIsEditable)
                            self.price_table.setItem(row_idx, 9, voice_alert_item)

                        if (self.alert_enabled[symbol] and
                                is_trading_now and
                                time_diff_seconds > tolerance_seconds):
                            alert_item.setText("開啟")
                            alert_item.setBackground(Qt.GlobalColor.green)
                            if voice_alert_item.text() != "已播放":
                                wav_path = self.wav_paths.get(symbol, "")
                                if wav_path and Path(wav_path).exists():
                                    try:
                                        threading.Thread(target=playsound, args=(wav_path,), daemon=True).start()
                                        self.log_play_event(symbol, current_time.strftime('%Y-%m-%d %H:%M:%S'))
                                        voice_alert_item.setText("已播放")
                                    except Exception as e:
                                        print(f"播放 {symbol} 的 WAV 失敗: {str(e)}")
                        else:
                            alert_item.setText("關閉")
                            alert_item.setBackground(QColor(255, 200, 200))

                        if (self.alert_enabled[symbol] and
                                is_trading_now and
                                time_diff_seconds < tolerance_seconds and
                                voice_alert_item.text() == "已播放"):
                            voice_alert_item.setText("")
                            wav_path_resume = self.wav_paths_resume.get(symbol, "")
                            if wav_path_resume and Path(wav_path_resume).exists():
                                try:
                                    threading.Thread(target=playsound, args=(wav_path_resume,), daemon=True).start()
                                    self.log_play_event(symbol, current_time.strftime('%Y-%m-%d %H:%M:%S'))
                                except Exception as e:
                                    print(f"播放 {symbol} 的恢復 WAV 失敗: {str(e)}")

                    self.status_label.setText(
                        f"最後更新: {current_time.strftime('%Y-%m-%d %H:%M:%S')} | 狀態: 運行中"
                    )
                    if self.auto_fit_enabled:
                        self.resize_table_columns()

        except Exception as e:
            print(f"顯示更新錯誤: {str(e)}")

    def monitor_prices(self):
        while self.running:
            with self.csv_lock:
                with open(self.csv_file, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Symbol", "Time", "Bid", "Ask"])

                    for symbol in self.symbols:
                        try:
                            mt5.symbol_select(symbol, True)
                            tick = mt5.symbol_info_tick(symbol)
                            if tick is not None:
                                now = datetime.datetime.now(tz=self.tz)
                                tick_time = datetime.datetime.fromtimestamp(tick.time, tz=self.tz)
                                if tick_time > now:
                                    tick_time = now
                                current_bid = tick.bid
                                current_ask = tick.ask

                                writer.writerow(
                                    [symbol, tick_time.strftime('%Y-%m-%d %H:%M:%S'), current_bid, current_ask])
                                self.last_prices[symbol] = {'bid': current_bid, 'ask': current_ask}

                                self.check_alert_conditions(symbol, current_bid, current_ask)

                        except Exception as e:
                            print(f"處理 {symbol} 時出錯: {str(e)}")

            time.sleep(0.1)

    def check_alert_conditions(self, symbol, bid, ask):
        config_file = "alert_config.csv"
        if Path(config_file).exists():
            df = pd.read_csv(config_file)
            symbol_config = df[df['symbol'] == symbol]
            if not symbol_config.empty and symbol_config['condition'].iloc[0] == "ON":
                wav_path = symbol_config['wav_path'].iloc[0]
                if Path(wav_path).exists():
                    threading.Thread(target=playsound, args=(wav_path,), daemon=True).start()
                    self.log_play_event(symbol, datetime.datetime.now(tz=self.tz).strftime('%Y-%m-%d %H:%M:%S'))

    def log_play_event(self, symbol, timestamp):
        with open("play_log.txt", "a") as f:
            f.write(f"{timestamp},{symbol},1\n")

    def is_trading(self, row, current_time):
        symbol = self.schedule_table.item(row, 0).text()
        start_day = self.schedule_table.cellWidget(row, 2).currentText()
        start_time_edit = self.schedule_table.cellWidget(row, 3)
        start_time_str = start_time_edit.time().toString("HH:mm:ss") if start_time_edit else "00:00:00"

        end_day = self.schedule_table.cellWidget(row, 4).currentText()
        end_time_edit = self.schedule_table.cellWidget(row, 5)
        end_time_str = end_time_edit.time().toString("HH:mm:ss") if end_time_edit else "00:00:00"

        day_map = {"星期一": 0, "星期二": 1, "星期三": 2, "星期四": 3,
                   "星期五": 4, "星期六": 5, "星期日": 6}

        current_weekday = current_time.weekday()
        week_start = current_time - datetime.timedelta(days=current_weekday)

        start_day_num = day_map[start_day]
        end_day_num = day_map[end_day]

        start_time = datetime.datetime.strptime(start_time_str, '%H:%M:%S').time()
        end_time = datetime.datetime.strptime(end_time_str, '%H:%M:%S').time()

        start_datetime = datetime.datetime.combine(
            week_start.date() + datetime.timedelta(days=start_day_num),
            start_time,
            tzinfo=self.tz
        )
        end_datetime = datetime.datetime.combine(
            week_start.date() + datetime.timedelta(days=end_day_num),
            end_time,
            tzinfo=self.tz
        )

        if end_day_num < start_day_num or (end_day_num == start_day_num and end_time < start_time):
            end_datetime += datetime.timedelta(days=7)

        within_trading_hours = start_datetime <= current_time <= end_datetime

        break_pairs = [(6, 7), (8, 9), (10, 11)]
        in_break = False

        current_day_num = current_time.weekday()
        current_day_start = current_time.replace(hour=0, minute=0, second=0, microsecond=0)

        for i, (start_col, end_col) in enumerate(break_pairs, 1):
            start_edit = self.schedule_table.cellWidget(row, start_col)
            end_edit = self.schedule_table.cellWidget(row, end_col)

            break_start_str = start_edit.time().toString("HH:mm:ss") if start_edit else "00:00:00"
            break_end_str = end_edit.time().toString("HH:mm:ss") if end_edit else "00:00:00"

            if break_start_str == "00:00:00" and break_end_str == "00:00:00":
                continue

            break_start_time = datetime.datetime.strptime(break_start_str, '%H:%M:%S').time()
            break_end_time = datetime.datetime.strptime(break_end_str, '%H:%M:%S').time()

            break_start = datetime.datetime.combine(
                current_day_start.date(),
                break_start_time,
                tzinfo=self.tz
            )
            break_end = datetime.datetime.combine(
                current_day_start.date(),
                break_end_time,
                tzinfo=self.tz
            )

            if break_end < break_start:
                break_end += datetime.timedelta(days=1)

            if break_start <= current_time <= break_end:
                in_break = True
                break

        return within_trading_hours and not in_break

    def start_monitoring(self):
        if not self.running:
            if not mt5.initialize():
                print("無法初始化 MT5")
                QMessageBox.critical(self, "錯誤", "無法初始化 MT5")
                return

            self.running = True
            self.refresh_timer.start(1000)

            self.price_thread = threading.Thread(target=self.monitor_prices)
            self.price_thread.daemon = True
            self.price_thread.start()
            self.status_label.setText(f"最後更新: 尚未更新 | 狀態: 運行中")
            print("監控已開始")

    def stop_monitoring(self):
        self.running = False
        self.refresh_timer.stop()
        self.status_label.setText(f"最後更新: 已停止 | 狀態: 已停止")
        print("監控已停止")

    def closeEvent(self, event):
        self.save_config()
        self.running = False
        self.refresh_timer.stop()
        mt5.shutdown()
        event.accept()

    def showEvent(self, event):
        super().showEvent(event)
        self.resize_table_columns()
        self.auto_fit_action.setChecked(self.auto_fit_enabled)

    def resize_table_columns(self):
        if self.auto_fit_enabled:
            self.price_table.resizeColumnsToContents()
            self.price_table.resizeRowsToContents()
            self.schedule_table.resizeColumnsToContents()
            self.schedule_table.resizeRowsToContents()
            self.product_table.resizeColumnsToContents()
            self.product_table.resizeRowsToContents()

    def toggle_auto_fit(self, checked):
        self.auto_fit_enabled = checked
        if checked:
            self.resize_table_columns()

    def create_toolbar(self):
        toolbar = QToolBar("主工具欄")
        self.addToolBar(toolbar)

        save_menu = QMenu("保存", self)
        save_params_action = QAction("保存參數", self)
        save_params_action.triggered.connect(self.save_parameters)
        save_menu.addAction(save_params_action)

        save_products_action = QAction("保存產品", self)
        save_products_action.triggered.connect(self.save_product_list)
        save_menu.addAction(save_products_action)

        self.auto_fit_action = QAction("自動調整", self)
        self.auto_fit_action.setCheckable(True)
        self.auto_fit_action.setChecked(True)
        self.auto_fit_action.triggered.connect(self.toggle_auto_fit)

        toggle_all_alerts_action = QAction("切換所有警報", self)
        toggle_all_alerts_action.triggered.connect(self.toggle_all_alerts)

        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)

        toolbar.addAction(save_menu.menuAction())
        toolbar.addSeparator()
        toolbar.addAction(self.auto_fit_action)
        toolbar.addAction(toggle_all_alerts_action)
        toolbar.addAction(exit_action)

    def toggle_all_alerts(self):
        all_checked = all(self.alert_enabled[symbol] for symbol in self.symbols)
        new_state = not all_checked

        for i, symbol in enumerate(self.symbols):
            self.alert_enabled[symbol] = new_state
            checkbox = self.price_table.cellWidget(i, 0)
            if checkbox:
                checkbox.setChecked(new_state)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PriceMonitor()
    window.show()
    sys.exit(app.exec())
