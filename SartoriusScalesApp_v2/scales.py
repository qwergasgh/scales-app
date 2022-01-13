from PyQt4 import QtCore
from datetime import datetime
import serial, re, logging, time
from settings import Settings

config = Settings()

class ScalesThread(QtCore.QThread):
    signal_scales = QtCore.pyqtSignal(str, int)
    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
        logging.info('Create thread scales scan port')
        self.running = False
        self.count = 0

    def convert_bytesize(self, value):
        value = int(value)
        if value == 5:
            return serial.FIVEBITS
        if value == 6:
            return serial.SIXBITS
        if value == 7:
            return serial.SEVENBITS
        if value == 8:
            return serial.EIGHTBITS

    def convert_parity(self, value):
        value = str(value)
        if value == "NONE":
            return serial.PARITY_NONE
        if value == "ODD":
            return serial.PARITY_ODD
        if value == "EVEN":
            return serial.PARITY_EVEN
        if value == "MARK":
            return serial.PARITY_MARK
        if value == "SPACE":
            return serial.PARITY_SPACE

    def convert_stopbit(self, value):
        value = float(value)
        if value == 1:
            return serial.STOPBITS_ONE
        if value == 1.5:
            return serial.STOPBITS_ONE_POINT_FIVE
        if value == 2:
            return serial.STOPBITS_TWO

    def write_values(self, line, count):
        logging.info('Parsing signal from scales')
        line_scales = line.strip().decode()

        if line_scales == "":
            return

        weight = ""
        # found = re.search(r'^N\s+[+-]\s*(\d+\.\d+) g', line_scales)
        found = re.search(r'^N\s+[+-]\s*(\d+\.\d+)($|( g))', line_scales)
        if found is not None:
            weight = str(found.group(1))

        if weight == "":
            logging.warning('Signal from scales -> weight is empty')
            return
        self.signal_scales.emit(weight, count)

    def run(self):
        self.running = True
        self.ser = serial.Serial(config.device, config.connection_speed, timeout=2, 
                           bytesize=self.convert_bytesize(config.data_bits), 
                           parity=self.convert_parity(config.parity_check), 
                           stopbits=self.convert_stopbit(config.stop_bit), 
                           rtscts=False, xonxoff=True, dsrdtr=False )
        try:
            while self.count < int(config.number_of_weighings):
                time.sleep(int(config.weighing_frequency))
                self.ser.write(b'\x1B\x50')
                line = self.ser.readline()
                self.write_values(line, self.count)
                self.count += 1
        except KeyboardInterrupt:
            exit()
