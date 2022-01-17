from exel_manager import ExelManager
from PyQt4 import QtGui, QtCore
from scales import ScalesThread
from main import config
import logging
# from settings import Settings



#config = Settings()

class ActiveBooksComboBox(QtGui.QComboBox):
    popupShown = QtCore.pyqtSignal()

    def showPopup(self):
        self.popupShown.emit()
        super(ActiveBooksComboBox, self).showPopup()


class AppWindow(QtGui.QWidget):
    def __init__(self, parent=None):
        logging.info("Create central widget")
        QtGui.QWidget.__init__(self, parent)

        bytes_size = ["5", "6", "7", "8"]
        parity = ["NONE", "ODD", "EVEN", "MARK", "SPACE"]
        stopbit = ["1", "1.5", "2"]

        self.exel_manager = ExelManager()

        form = QtGui.QFormLayout()

        # status work
        self.label_status_work = QtGui.QLabel("<b>Остановлено</b>")
        self.label_status_work.setAlignment(QtCore.Qt.AlignCenter)
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

        # number of weighings
        self.lineedit_number_of_weighings = QtGui.QLineEdit()
        self.lineedit_number_of_weighings.setText(config.number_of_weighings)
        form.addRow("Количество взвешиваний:", self.lineedit_number_of_weighings)

        # weighing frequency
        self.lineedit_weighing_frequency = QtGui.QLineEdit()
        self.lineedit_weighing_frequency.setText(config.weighing_frequency)
        form.addRow("Периодичность взвешиваний:", self.lineedit_weighing_frequency)

        # active book
        self.combobox_active_exel = ActiveBooksComboBox(self) # QtGui.QComboBox()
        count = 0
        for name in self.exel_manager.get_books():
            self.combobox_active_exel.insertItem(count, name)
            count += 1
        form.addRow("Активный EXCEL файл:", self.combobox_active_exel)
        self.exel_manager.set_active_book(self.combobox_active_exel.currentText())

        # buttons
        self.button_apply_settings= QtGui.QPushButton("Применить")
        self.button_start_scan = QtGui.QPushButton("Запуск")
        self.button_stop_scan = QtGui.QPushButton("Остановить")
        self.button_stop_scan.setDisabled(False)

        # button apply
        form.addRow(self.button_apply_settings)
        # button start
        form.addRow(self.button_start_scan)
        # button stop
        form.addRow(self.button_stop_scan)
        self.button_stop_scan.setDisabled(True)

        self.setLayout(form)

        self.scales = ScalesThread()

        self.button_apply_settings.clicked.connect(self.apply_settings)
        self.button_start_scan.clicked.connect(self.start)
        self.button_stop_scan.clicked.connect(self.stop)

        self.scales.signal_scales.connect(self.write_values_to_exel, QtCore.Qt.QueuedConnection)
        self.combobox_active_exel.popupShown.connect(self.update_active_books)
	
    def update_active_books(self):
        active_books = self.exel_manager.update_books()
        if len(active_books) == 0:
            logging.info("Not found excel")
        active_book = self.combobox_active_exel.currentText()
        self.combobox_active_exel.clear()
        count = 0
        for name in active_books:
            self.combobox_active_exel.insertItem(count, name)
            count += 1
        if active_book in active_books:
            self.exel_manager.set_active_book(active_book)
        else:
            self.exel_manager.set_active_book(self.combobox_active_exel.currentText())
    
    def start(self):
        if not self.scales.isRunning():
            self.button_start_scan.setDisabled(True)
            self.button_stop_scan.setDisabled(False)
            self.button_apply_settings.setDisabled(True)
            self.scales.count = 0
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
        self.button_apply_settings.setDisabled(False)
        self.label_status_work.setText("<b>Остановлено</b>")
        logging.info('Stop scan port')

    def apply_settings(self):
        logging.info('Aplly settings')
        self.exel_manager.set_active_book(self.combobox_active_exel.currentText())
        config.update({"device": self.lineedit_device.text(), "connection_speed": self.lineedit_connection_speed.text(),
                       "data_bits": self.combobox_data_bits.currentText(), "parity_check": self.combobox_parity_check.currentText(),
                       "stop_bit": self.combobox_stop_bit.currentText(), "number_of_weighings": self.lineedit_number_of_weighings.text(), 
                       "weighing_frequency": self.lineedit_weighing_frequency.text()})

    def write_values_to_exel(self, weight, count):
        logging.info(str(count + 1) + "/" + config.number_of_weighings)
        if count == 0:
            self.exel_manager.set_parametr_write()
        logging.info("Signal from scales to exel (weight = " + weight + ")")
        self.exel_manager.write_values(weight, count)
        if (count + 1) == int(config.number_of_weighings):
            self.stop()
            self.scales.count = 0
            self.exel_manager.clear_parametrs()

    def closeEvent(self, event):
        config.save()
        self.hide()
        self.scales.running = False
        self.scales.ser.close()
        self.scales.wait(500)
        event.accept()
        logging.info('Close app')
        