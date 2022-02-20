from excel_manager import ExcelManager
from PyQt4 import QtGui, QtCore
from scales import ScalesThread
from main import config
import logging


class ActiveBooksComboBox(QtGui.QComboBox):
    popupShown = QtCore.pyqtSignal()

    def showPopup(self):
        self.popupShown.emit()
        super(ActiveBooksComboBox, self).showPopup()


class AppWindow(QtGui.QWidget):
    def __init__(self, parent=None):
        logging.info("Create central widget")
        QtGui.QWidget.__init__(self, parent)

        abc = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
               "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
        bytes_size = ("5", "6", "7", "8")
        parity = ("NONE", "ODD", "EVEN", "MARK", "SPACE")
        stopbit = ("1", "1.5", "2")
        self.modes = {"Масса в сумме": 0, "Масса корточки": 1}

        self.excel_manager = ExcelManager()

        form = QtGui.QFormLayout()

        # status work
        self.label_status_work = QtGui.QLabel("<b>Остановлено</b>")
        form.addRow("<b>Статус работы:</b>", self.label_status_work)

        # device
        self.lineedit_device = QtGui.QLineEdit()
        self.lineedit_device.setText(config.device)
        form.addRow("Устройство:", self.lineedit_device)

        # speed
        self.lineedit_connection_speed = QtGui.QLineEdit()
        self.lineedit_connection_speed.setText(config.connection_speed)
        form.addRow("Скорость соединения:", self.lineedit_connection_speed)

        # bits
        self.combobox_data_bits = QtGui.QComboBox()
        count = 0
        for i in bytes_size:
            self.combobox_data_bits.insertItem(count, i)
            count += 1
        self.combobox_data_bits.setCurrentIndex(bytes_size.index(config.data_bits))
        form.addRow("Количество битов данных:", self.combobox_data_bits)

        # parity
        self.combobox_parity_check = QtGui.QComboBox()
        count = 0
        for i in parity:
            self.combobox_parity_check.insertItem(count, i)
            count += 1
        self.combobox_parity_check.setCurrentIndex(parity.index(config.parity_check))
        form.addRow("Проверка четности:", self.combobox_parity_check)

        # stop bit
        self.combobox_stop_bit = QtGui.QComboBox()
        count = 0
        for i in stopbit:
            self.combobox_stop_bit.insertItem(count, i)
            count += 1
        self.combobox_stop_bit.setCurrentIndex(stopbit.index(config.stop_bit))
        form.addRow("Стоп бит:", self.combobox_stop_bit)

        hbox_weight = QtGui.QHBoxLayout()
        # mass sum
        groupbox_weight_sum = QtGui.QGroupBox("Масса в сумме")
        widget_data_sum = QtGui.QWidget()
        hbox_data_sum = QtGui.QHBoxLayout()
        label_data_sum = QtGui.QLabel("Дата:")
        self.combobox_date_sum = QtGui.QComboBox()
        count = 0
        for i in abc:
            self.combobox_date_sum.insertItem(count, i)
            count += 1
        self.combobox_date_sum.setCurrentIndex(abc.index(config.date_sum))
        hbox_data_sum.addWidget(label_data_sum)
        hbox_data_sum.addWidget(self.combobox_date_sum)
        widget_data_sum.setLayout(hbox_data_sum)

        widget_weight_sum = QtGui.QWidget()
        hbox_weight_sum = QtGui.QHBoxLayout()
        label_weight_sum = QtGui.QLabel("Масса")
        self.combobox_weight_sum = QtGui.QComboBox()
        count = 0
        for i in abc:
            self.combobox_weight_sum.insertItem(count, i)
            count += 1
        self.combobox_weight_sum.setCurrentIndex(abc.index(config.weight_sum))
        hbox_weight_sum.addWidget(label_weight_sum)
        hbox_weight_sum.addWidget(self.combobox_weight_sum)
        widget_weight_sum.setLayout(hbox_weight_sum)

        main_vbox_weight_sum = QtGui.QVBoxLayout()
        main_vbox_weight_sum.addWidget(widget_data_sum)
        main_vbox_weight_sum.addWidget(widget_weight_sum)
        groupbox_weight_sum.setLayout(main_vbox_weight_sum)

        # mass kort
        groupbox_weight_kort = QtGui.QGroupBox("Масса корточки")
        widget_data_kort = QtGui.QWidget()
        hbox_data_kort = QtGui.QHBoxLayout()
        label_data_kort = QtGui.QLabel("Дата:")
        self.combobox_date_kort = QtGui.QComboBox()
        count = 0
        for i in abc:
            self.combobox_date_kort.insertItem(count, i)
            count += 1
        self.combobox_date_kort.setCurrentIndex(abc.index(config.date_kort))
        hbox_data_kort.addWidget(label_data_kort)
        hbox_data_kort.addWidget(self.combobox_date_kort)
        widget_data_kort.setLayout(hbox_data_kort)

        widget_weight_kort = QtGui.QWidget()
        hbox_weight_kort = QtGui.QHBoxLayout()
        label_weight_kort = QtGui.QLabel("Масса")
        self.combobox_weight_kort = QtGui.QComboBox()
        count = 0
        for i in abc:
            self.combobox_weight_kort.insertItem(count, i)
            count += 1
        self.combobox_weight_kort.setCurrentIndex(abc.index(config.weight_kort))
        hbox_weight_kort.addWidget(label_weight_kort)
        hbox_weight_kort.addWidget(self.combobox_weight_kort)
        widget_weight_kort.setLayout(hbox_weight_kort)

        main_vbox_weight_kort = QtGui.QVBoxLayout()
        main_vbox_weight_kort.addWidget(widget_data_kort)
        main_vbox_weight_kort.addWidget(widget_weight_kort)
        groupbox_weight_kort.setLayout(main_vbox_weight_kort)

        hbox_weight = QtGui.QHBoxLayout()
        hbox_weight.addWidget(groupbox_weight_sum)
        hbox_weight.addWidget(groupbox_weight_kort)
        form.addRow(hbox_weight)

        # active book
        self.combobox_active_excel = ActiveBooksComboBox(self)  # QtGui.QComboBox()
        count = 0
        for name in self.excel_manager.get_books():
            self.combobox_active_excel.insertItem(count, name)
            count += 1
        form.addRow("Активный EXCEL файл:", self.combobox_active_excel)
        self.excel_manager.set_active_book(self.combobox_active_excel.currentText())

        button_apply_settings = QtGui.QPushButton("Применить")
        self.button_start_scan = QtGui.QPushButton("Запуск")
        self.button_stop_scan = QtGui.QPushButton("Остановить")
        self.button_stop_scan.setDisabled(False)

        # mode work
        self.combobox_work = QtGui.QComboBox()
        count = 0
        for i in self.modes.keys():
            self.combobox_work.insertItem(count, i)
            count += 1
        form.addRow("Режим работы", self.combobox_work)

        # button apply
        form.addRow(button_apply_settings)
        # button start
        form.addRow(self.button_start_scan)
        # button stop
        form.addRow(self.button_stop_scan)
        self.button_stop_scan.setDisabled(True)

        self.setLayout(form)

        self.scales = ScalesThread()

        button_apply_settings.clicked.connect(self.apply_settings)
        self.button_start_scan.clicked.connect(self.start)
        self.button_stop_scan.clicked.connect(self.stop)

        self.scales.signal_scales.connect(self.write_values_to_exel, QtCore.Qt.QueuedConnection)
        self.combobox_active_excel.popupShown.connect(self.update_active_books)

    def update_active_books(self):
        active_books = self.excel_manager.update_books()
        if len(active_books) == 0:
            logging.info("Not found excel")
        active_book = self.combobox_active_excel.currentText()
        self.combobox_active_excel.clear()
        count = 0
        for name in active_books:
            self.combobox_active_excel.insertItem(count, name)
            count += 1
        if active_book in active_books:
            self.excel_manager.set_active_book(active_book)
        else:
            self.excel_manager.set_active_book(self.combobox_active_excel.currentText())

    def start(self):
        if not self.scales.isRunning():
            self.button_start_scan.setDisabled(True)
            self.button_stop_scan.setDisabled(False)
            self.scales.start()
            self.label_status_work.setText("<b>В работе</b>")
            logging.info('Start scan port')

    def stop(self):
        self.scales.running = False
        self.scales.terminate()
        self.scales.quit()
        self.scales.ser.close()
        self.button_start_scan.setDisabled(False)
        self.button_stop_scan.setDisabled(True)
        self.label_status_work.setText("<b>Остановлено</b>")
        logging.info('Stop scan port')

    def apply_settings(self):
        logging.info('Aplly settings')
        self.excel_manager.set_active_book(self.combobox_active_excel.currentText())
        config.update({"device": self.lineedit_device.text(), "connection_speed": self.lineedit_connection_speed.text(),
                       "data_bits": self.combobox_data_bits.currentText(),
                       "parity_check": self.combobox_parity_check.currentText(),
                       "stop_bit": self.combobox_stop_bit.currentText(),
                       "date_sum": self.combobox_date_sum.currentText(),
                       "weight_sum": self.combobox_weight_sum.currentText(),
                       "date_kort": self.combobox_date_kort.currentText(),
                       "weight_kort": self.combobox_weight_kort.currentText(),
                       "mode": str(self.modes[self.combobox_work.currentText()])})

    def write_values_to_exel(self, date, weight):
        if int(config.mode) == 0:
            self.excel_manager.write_values(date, weight, config.date_sum, config.weight_sum)
        if int(config.mode) == 1:
            self.excel_manager.write_values(date, weight, config.date_kort, config.weight_kort)

    def closeEvent(self, event):
        config.save()
        self.hide()
        self.scales.running = False
        self.scales.wait(500)
        event.accept()
        logging.info('Close app')
