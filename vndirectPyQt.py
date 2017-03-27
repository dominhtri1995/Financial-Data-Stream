from selenium import webdriver
import pandas as pd
import threading
from threading import Semaphore
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import encodings.idna
import os
import sqlite3
import logging
import datetime
import time


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def get_timestamp():
    ts = time.time()
    return datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')


def get_statement(type, ticker, excel):
    lock.acquire()
    for i in range(0, 2):
        try:
            driver = webdriver.Chrome(executable_path=resource_path("./chromedriver"))

            driver.set_window_position(5000, 0)
            if (type == "bs"):
                url = "https://www.vndirect.com.vn/portal/bang-can-doi-ke-toan/" + ticker + ".shtml"
            elif (type == "ic"):
                url = "https://www.vndirect.com.vn/portal/bao-cao-ket-qua-kinh-doanh/" + ticker + ".shtml"
            elif (type == "cf"):
                url = "https://www.vndirect.com.vn/portal/bao-cao-luu-chuyen-tien-te/" + ticker + ".shtml"
            driver.get(url)

            year_mode = driver.find_element_by_name("searchObject.fiscalQuarter")
            all_options = year_mode.find_elements_by_tag_name("option")
            for option in all_options:
                if (option.get_attribute("value") == "IN_YEAR"):
                    option.click()

            xem = driver.find_element_by_xpath("//input[@class='iButton autoHeight']")
            xem.click()

            html = driver.page_source
            df = pd.read_html(html)
            df = df[1]
            # print(df)
            df = df.set_index(0)
            df.to_excel(excel, type)
            excel.save()
        except:
            e = sys.exc_info()[0]
            e1 = sys.exc_info()[1]
            logging.info(get_timestamp() + str(e) + str(e1))
            try:
                driver.close()
            except:
                pass
            continue
        driver.close()
        break

    lock.release()


def get_data(ticker, bs, ic, cf):
    ticker = ticker.upper()
    tickers = ticker.split(",")
    # determine if application is a script file or frozen exe

    for ticker in tickers:

        excelwriter = pd.ExcelWriter(os.path.join(dir, ticker + ".xlsx"))
        logging.info(get_timestamp() + os.path.join(dir, ticker + ".xlsx"))
        if (bs == 1):
            # text.insert(INSERT,"Getting Balance Sheet")
            thread1 = myThread(1, "bs", ticker, excelwriter)
            thread1.start()
            threads.append(thread1)
        if (ic == 1):
            thread2 = myThread(1, "ic", ticker, excelwriter)
            thread2.start()
            threads.append(thread2)
        if (cf == 1):
            thread3 = myThread(1, "cf", ticker, excelwriter)
            thread3.start()
            threads.append(thread3)


class myThread(threading.Thread):
    def __init__(self, threadID, name, ticker, excelwriter):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.name = name
        self.ticker = ticker
        self.excelwriter = excelwriter

    def run(self):
        logging.info("Starting " + self.name)
        get_statement(self.name, self.ticker, self.excelwriter)


###********** MAIN PROGRAM *************##
# global variable go here
lock = Semaphore(3)
threads = []
dir = "./"
if getattr(sys, 'frozen', False):
    dir = os.path.dirname(sys.executable)
elif __file__:
    dir = os.path.dirname(__file__)
else:
    logging.info("detect app")
    dir = os.getcwd()


######### GUI go here #########
class GUI(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        overall_layout = QtWidgets.QVBoxLayout()
        main_layout = QtWidgets.QVBoxLayout()
        malayout = QtWidgets.QVBoxLayout()
        hboxboth = QtWidgets.QHBoxLayout()
        hboxboth.setAlignment(QtCore.Qt.AlignCenter)

        #### Tabs
        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        # self.tabs.resize(300, 200)

        # Add tabs
        self.tabs.addTab(self.tab1, "Financial Statement")
        self.tabs.addTab(self.tab2, "M&&A")

        ### Form for ticker ####
        label = QLabel()
        label.setText("Ticker")
        tickerEdit = QLineEdit()
        formlayout = QtWidgets.QFormLayout()
        formlayout.addRow(label, tickerEdit)
        formlayout.setSpacing(5)
        ##### Type #####
        bs = QtWidgets.QCheckBox('Balance sheet', self)
        ic = QtWidgets.QCheckBox('Income Statement', self)
        cf = QtWidgets.QCheckBox('Cash Flow', self)
        bs.toggle()
        ic.toggle()
        cf.toggle()

        vboxtype = QtWidgets.QVBoxLayout()
        vboxtype.addWidget(bs)
        vboxtype.addWidget(ic)
        vboxtype.addWidget(cf)
        vboxtype.setSpacing(10)

        hboxboth.addLayout(formlayout)
        hboxboth.addLayout(vboxtype)
        hboxboth.setSpacing(10)
        main_layout.addItem(hboxboth)

        ### Download button
        btn = QtWidgets.QPushButton("Download", self)
        btn.clicked.connect(lambda: get_data(tickerEdit.text(), bs.isChecked(), ic.isChecked(), cf.isChecked()))
        hbox = QtWidgets.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(btn)
        hbox.addStretch(1)

        vbox = QtWidgets.QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox)
        vbox.addStretch(1)
        main_layout.addItem(vbox)

        self.tab1.setLayout(main_layout)

        malayout.addWidget(MA())
        malayout.setContentsMargins(5, 5, 5, 5)
        self.tab2.setLayout(malayout)
        overall_layout.addWidget(self.tabs)

        self.setLayout(overall_layout)
        self.setGeometry(50, 50, 1100, 1100)
        self.setWindowTitle('Whirlpool-data')
        self.show()

    def closeEvent(self, event):

        reply = QtWidgets.QMessageBox.question(self, 'Message',
                                               "Are you sure to quit?", QtWidgets.QMessageBox.Yes |
                                               QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)

        if reply == QtWidgets.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


class MA(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        self.createTable()

        # Add box layout, add table to box layout and add box layout to widget
        self.layout = QVBoxLayout()
        self.layout.addItem(self.hbox)
        self.layout.addWidget(self.tableWidget)
        self.setLayout(self.layout)

    def createTable(self):
        self.hbox = QtWidgets.QHBoxLayout()
        self.filter = QtWidgets.QPushButton('', self)
        self.filter.setIcon(QtGui.QIcon(resource_path('./image/filter.png')))
        self.dialogFilter = FilterDialog()
        self.connect_filter_dilaog(self.dialogFilter)

        self.filter.clicked.connect(self.filter_window)

        self.toexcel = QtWidgets.QPushButton('', self)
        self.toexcel.setIcon(QtGui.QIcon(resource_path('./image/excel.png')))
        self.toexcel.clicked.connect(self.to_excel)

        # Clear filter
        self.clearfilter = QPushButton('', self)
        self.clearfilter.setIcon(QtGui.QIcon(resource_path('./image/reset.png')))
        self.clearfilter.clicked.connect(self.clear_filter)

        self.hbox.addStretch(1)
        self.hbox.addWidget(self.clearfilter)
        self.hbox.addWidget(self.filter)
        self.hbox.addWidget(self.toexcel)
        self.hbox.setSpacing(10)

        # Table data
        self.tableWidget = QTableWidget()
        self.tableWidget.setColumnCount(11)
        self.tableWidget.setHorizontalHeaderLabels(
            ["Date", "Deal size(m)", "Status", "Target Name", "Target Industry", "Acquirer", \
             "Acquirer Industry", "Percent Deal", "Target Nation", "Transaction", "Net Income Multiple"])
        self.tableWidget.setSortingEnabled(True)
        # get data from database
        self.conn = sqlite3.connect(resource_path('./testDB.db'))
        self.sort_statement = "SELECT ANN_DATE,SIZE,STATUS,TARGET_NAME,TARGET_INDUSTRY,ACQUIRER_NAME \
							  , ACQUIRER_INDUSTRY,PERCENT,TARGET_NATION,FORM,NI_MULTIPLE \
								FROM TRDATA"
        self.execute_sql_statement()

    def execute_sql_statement(self):
        cursor = self.conn.execute(self.sort_statement)
        rows = cursor.fetchall()
        self.tableWidget.setRowCount(len(rows))
        r = 0
        for row in rows:
            for c in range(0, 11, 1):
                self.tableWidget.setItem(r, c, QTableWidgetItem(str(row[c])))
            r += 1

    def to_excel(self):
        self.conn = sqlite3.connect(resource_path('./testDB.db'))
        df = pd.read_sql_query(self.sort_statement, self.conn)
        writer = pd.ExcelWriter("export.xlsx")
        df.to_excel(writer, "Sheet 1")
        writer.save()

    def filter_window(self):
        self.dialogFilter.show()

    def clear_filter(self):
        self.sort_statement = "SELECT ANN_DATE,SIZE,STATUS,TARGET_NAME,TARGET_INDUSTRY,ACQUIRER_NAME \
							  , ACQUIRER_INDUSTRY,PERCENT,TARGET_NATION,FORM,NI_MULTIPLE \
								FROM TRDATA"
        self.execute_sql_statement()

    def connect_filter_dilaog(self, filterDialog):
        filterDialog.containsTargetName.connect(self.filter_table)

    @QtCore.pyqtSlot('QString', 'QString', 'QString', 'QString', 'QString', 'QString', 'QString', 'QString', 'QDate')
    def filter_table(self, target_name, acquirer_name, deaL_size_sign, deal_size, deal_status \
                     , target_industry, acquirer_indsutry, date_sign, date):

        logging.info("filter: change table with tn: " + target_name)
        self.sort_statement = "SELECT ANN_DATE,SIZE,STATUS,TARGET_NAME,TARGET_INDUSTRY,ACQUIRER_NAME \
							  , ACQUIRER_INDUSTRY,PERCENT,TARGET_NATION,FORM,NI_MULTIPLE \
								FROM TRDATA"

        if (target_name != ""):
            self.sort_statement = self.sort_statement + " WHERE TARGET_NAME LIKE '%" + target_name + "%'"
        if (acquirer_name != ""):
            if ("WHERE" in self.sort_statement):
                self.sort_statement = self.sort_statement + " AND ACQUIRER_NAME LIKE '%" + acquirer_name + "%'"
            else:
                self.sort_statement = self.sort_statement + " WHERE ACQUIRER_NAME LIKE '%" + acquirer_name + "%'"
        if (deal_size != ""):
            if ("WHERE" in self.sort_statement):
                self.sort_statement = self.sort_statement + " AND SIZE " + deaL_size_sign + " " + deal_size
            else:
                self.sort_statement = self.sort_statement + " WHERE SIZE " + deaL_size_sign + " " + deal_size

        if (deal_status != "None"):
            if ("WHERE" in self.sort_statement):
                self.sort_statement = self.sort_statement + " AND STATUS LIKE '%" + deal_status + "%'"
            else:
                self.sort_statement = self.sort_statement + " WHERE STATUS LIKE '%" + deal_status + "%'"

        if (target_industry != ""):
            if ("WHERE" in self.sort_statement):
                self.sort_statement = self.sort_statement + " AND TARGET_INDUSTRY LIKE '%" + target_industry + "%'"
            else:
                self.sort_statement = self.sort_statement + " WHERE TARGET_INDUSTRY LIKE '%" + target_industry + "%'"

        if (acquirer_indsutry != ""):
            if ("WHERE" in self.sort_statement):
                self.sort_statement = self.sort_statement + " AND ACQUIRER_INDUSTRY LIKE '%" + acquirer_indsutry + "%'"
            else:
                self.sort_statement = self.sort_statement + " WHERE ACQUIRER_INDUSTRY LIKE '%" + acquirer_indsutry + "%'"

        if (date_sign != "None"):
            # date = "'"+str(date.year()) + "-" + str(date.month()) + "-" + str(date.day())+ "'"
            if (date_sign == "After"):
                date_sign = ">"
            elif (date_sign == "Before"):
                date_sign = "<"
            else:
                date_sign = "="

            if ("WHERE" in self.sort_statement):
                self.sort_statement = self.sort_statement + " AND ANN_DATE " + date_sign + " '" + date.toString(
                    format=Qt.ISODate) + "'"
            else:
                self.sort_statement = self.sort_statement + " WHERE ANN_DATE " + date_sign + " '" + date.toString(
                    format=Qt.ISODate) + "'"

        print(self.sort_statement)
        self.execute_sql_statement()


class FilterDialog(QDialog):
    containsTargetName = QtCore.pyqtSignal('QString', 'QString', 'QString', 'QString', 'QString', 'QString', 'QString',
                                           'QString', 'QDate')

    def __init__(self, parent=None):
        super(FilterDialog, self).__init__(parent)
        self.initUI()

    def initUI(self):
        self.edit_to_clear = []
        self.combo_to_clear = []

        self.mainlayout = QVBoxLayout()
        self.mainlayout.setContentsMargins(20, 20, 20, 20)
        label_target_name = QLabel()
        label_target_name.setText("Target Name (contains):")
        target_name_edit = QLineEdit()
        target_name_edit.setPlaceholderText("Keywords")

        label_acquirer_name = QLabel()
        label_acquirer_name.setText("Acquirer Name (contains):")
        acquirer_name_edit = QLineEdit()
        acquirer_name_edit.setPlaceholderText("Keywords")

        label_deal_size = QLabel()
        label_deal_size.setText("Deal Size :")
        deal_size_edit = QLineEdit()
        deal_size_edit.setPlaceholderText("Size in m$")

        deal_size_layout = QHBoxLayout()
        combo_deal_size = QComboBox(self)
        combo_deal_size.addItem(">")
        combo_deal_size.addItem("=")
        combo_deal_size.addItem("<")
        combo_deal_size.addItem(">=")
        combo_deal_size.addItem("<=")
        deal_size_layout.addWidget(combo_deal_size)
        deal_size_layout.addWidget(deal_size_edit)

        label_status = QLabel()
        label_status.setText("Deal Status :")
        combo_status = QComboBox(self)
        combo_status.addItem("None")
        combo_status.addItem("Completed")
        combo_status.addItem("Pending")
        combo_status.addItem("Rumor")
        combo_status.addItem("Seeking Buyer")
        combo_status.addItem("Unknown")
        combo_status.addItem("Withdrawn")

        label_target_industry = QLabel()
        label_target_industry.setText("Target Industry :")
        target_industry_edit = QLineEdit()
        target_industry_edit.setPlaceholderText("Keywords")

        label_acquirer_industry = QLabel()
        label_acquirer_industry.setText("Acquirer Industry :")
        acquirer_industry_edit = QLineEdit()
        acquirer_industry_edit.setPlaceholderText("Keywords")

        ####calender
        calendar_layout = QHBoxLayout()
        label_calendar = QLabel()
        label_calendar.setText("Date :")
        cal = QCalendarWidget()
        cal.setGridVisible(True)
        combo_calendar = QComboBox(self)
        combo_calendar.addItem("None")
        combo_calendar.addItem("After")
        combo_calendar.addItem("Equal")
        combo_calendar.addItem("Before")
        calendar_layout.addWidget(combo_calendar)
        calendar_layout.addWidget(cal)

        formlayout = QtWidgets.QFormLayout()
        formlayout.addRow(label_target_name, target_name_edit)
        formlayout.addRow(label_acquirer_name, acquirer_name_edit)
        formlayout.addRow(label_deal_size, deal_size_layout)
        formlayout.addRow(label_status, combo_status)
        formlayout.addRow(label_target_industry, target_industry_edit)
        formlayout.addRow(label_acquirer_industry, acquirer_industry_edit)
        formlayout.addRow(label_calendar, combo_calendar)
        formlayout.addRow(cal)
        formlayout.setSpacing(15)
        formlayout.setLabelAlignment(Qt.AlignLeft)

        self.edit_to_clear.extend((target_name_edit, acquirer_name_edit \
                                       , deal_size_edit, target_industry_edit, acquirer_industry_edit))

        self.combo_to_clear.extend((combo_status, combo_calendar, combo_deal_size))
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Reset | QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttonBox.button(QDialogButtonBox.Reset).clicked.connect(self.clear_form)
        self.buttonBox.button(QDialogButtonBox.Ok).clicked.connect(lambda: self.closeok(target_name_edit.text() \
                                                                                        , acquirer_name_edit.text() \
                                                                                        , combo_deal_size.currentText() \
                                                                                        , deal_size_edit.text() \
                                                                                        , combo_status.currentText() \
                                                                                        , target_industry_edit.text() \
                                                                                        , acquirer_industry_edit.text() \
                                                                                        , combo_calendar.currentText() \
                                                                                        , cal.selectedDate()))
        self.buttonBox.button(QDialogButtonBox.Cancel).clicked.connect(self.close)

        self.mainlayout.addItem(formlayout)
        self.mainlayout.addWidget(self.buttonBox)
        self.setLayout(self.mainlayout)

    def closeok(self, target_name, acquirer_name, deaL_size_sign, deal_size, deal_status \
                , target_industry, acquirer_indsutry, date_sign, date):
        logging.info("filterDialog applied")
        self.containsTargetName.emit(target_name, acquirer_name, deaL_size_sign, deal_size, deal_status \
                                     , target_industry, acquirer_indsutry, date_sign, date)
        # self.close()

    def clear_form(self):
        for edit in self.edit_to_clear:
            edit.clear()
        for combo in self.combo_to_clear:
            combo.setCurrentIndex(0)


logging.basicConfig(filename='process.log', level=logging.INFO)
app = QtWidgets.QApplication(sys.argv)
gui = GUI()
sys.exit(app.exec_())
