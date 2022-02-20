from logging.handlers import RotatingFileHandler
from win32com.client import GetObject
from settings import Settings
from PyQt4 import QtGui
import logging
import time
import sys

logging.basicConfig(handlers=[RotatingFileHandler(filename="logfile.log", mode='w',
                                                  maxBytes=50000, backupCount=4)],
                    level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
config = Settings()


def get_excel_process():
    time.sleep(3)
    process_excel = False
    WMI = GetObject('winmgmts:')
    processes = WMI.InstancesOf('Win32_Process')
    for i in [process.Properties_('Name').Value.lower() for process in processes]:
        if i.find('excel.exe') > -1:
            process_excel = True
            break
    return process_excel


def show_message(css, icon):
    msg = QtGui.QMessageBox()
    msg.setWindowTitle("ScalesApp")
    msg.setText("Проверка запуска Exсel")
    msg.setInformativeText("Откройте необходимый файл Excel и выберите активную ячейку")
    msg.setWindowIcon(icon)
    msg.setIcon(QtGui.QMessageBox.Warning)
    msg.setStandardButtons(QtGui.QMessageBox.Ok | QtGui.QMessageBox.Close)
    button_ok = msg.button(QtGui.QMessageBox.Ok)
    button_ok.setText("Запустить")
    button_close = msg.button(QtGui.QMessageBox.Close)
    button_close.setText("Закрыть")
    msg.setStyleSheet(css)
    ret = msg.exec_()
    if ret == QtGui.QMessageBox.Ok:
        if get_excel_process() is False:
            logging.info("Not found excel")
            show_message(css, icon)
    if ret == QtGui.QMessageBox.Close:
        exit()


def get_css():
    css = []
    try:
        logging.info("Import styles app")
        with open("style.css", "r") as file:
            for line in file:
                css.append(line)
    except:
        logging.warning("Error import style")
        return ""
    return "".join(css)


def create_app():
    logging.info('Open app')
    from widgets import AppWindow
    css = get_css()
    app = QtGui.QApplication(sys.argv)
    icon = QtGui.QIcon("icon.png")
    app.setWindowIcon(icon)
    if get_excel_process() is False:
        show_message(css, icon)
    w = AppWindow()
    w.setStyleSheet(css)
    w.resize(250, 150)
    w.move(300, 300)
    w.setWindowTitle("ScalesApp_v1")
    w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    create_app()
