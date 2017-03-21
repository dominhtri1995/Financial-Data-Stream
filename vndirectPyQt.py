from selenium import webdriver
import pandas as pd
import threading
from threading import Semaphore
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
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
            e1=sys.exc_info()[1]
            logging.info(get_timestamp() + str(e) +str(e1))
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

    @QtCore.pyqtSlot("QString")
    def filter_table(self, targetName):
        logging.info("filter: change table with tn: " + targetName)
        self.sort_statement = "SELECT ANN_DATE,SIZE,STATUS,TARGET_NAME,TARGET_INDUSTRY,ACQUIRER_NAME \
							  , ACQUIRER_INDUSTRY,PERCENT,TARGET_NATION,FORM,NI_MULTIPLE \
								FROM TRDATA" + " WHERE TARGET_NAME LIKE '%" + str(targetName) + "%'"
        self.execute_sql_statement()


class FilterDialog(QDialog):
    containsTargetName = QtCore.pyqtSignal('QString')

    def __init__(self, parent=None):
        super(FilterDialog, self).__init__(parent)
        self.initUI()
        self.filter = 0

    def initUI(self):
        self.mainlayout = QVBoxLayout()
        self.mainlayout.setContentsMargins(20, 20, 20, 20)
        label = QLabel()
        label.setText("Target Name (contains)")
        targetNameEdit = QLineEdit()
        targetNameEdit.setPlaceholderText("Keywords for Target name")
        formlayout = QtWidgets.QFormLayout()
        formlayout.addRow(label, targetNameEdit)
        formlayout.setSpacing(5)

        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttonBox.button(QDialogButtonBox.Ok).clicked.connect(lambda: self.closeok(targetNameEdit.text()))
        self.buttonBox.button(QDialogButtonBox.Cancel).clicked.connect(self.close)

        self.mainlayout.addItem(formlayout)
        self.mainlayout.addWidget(self.buttonBox)
        self.setLayout(self.mainlayout)

    def closeok(self, targetName):
        # self.filter =1
        logging.info("filterDialog applied")
        self.containsTargetName.emit(targetName)

    # self.close()


logging.basicConfig(filename='process.log', level=logging.INFO)
app = QtWidgets.QApplication(sys.argv)
gui = GUI()
sys.exit(app.exec_())
