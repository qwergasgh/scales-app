from logging.handlers import RotatingFileHandler
from PyQt4 import QtGui
import logging, sys


logging.basicConfig(handlers=[RotatingFileHandler(filename="logfile.log",
                     mode='w', maxBytes=50000, backupCount=4)], level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def get_exel_process():
    import time
    from win32com.client import GetObject
    time.sleep(3)
    process_exel = False
    WMI = GetObject('winmgmts:')
    processes = WMI.InstancesOf('Win32_Process')
    for i in [process.Properties_('Name').Value.lower() for process in processes]:
        if i.find('excel.exe') > -1:
            process_exel = True
            break
    return process_exel
    
def show_message(css, icon):
    msg = QtGui.QMessageBox()
    msg.setWindowTitle("SartoriusScalesApp")
    msg.setText("Проверка запуска Exсel")
    msg.setInformativeText("Откройте необходимый файл Excel и выберите активную ячейку")
    msg.setWindowIcon(icon)
    msg.setIcon(QtGui.QMessageBox.Warning)
    msg.setStandardButtons(QtGui.QMessageBox.Ok | QtGui.QMessageBox.Close)
    buutom_ok = msg.button(QtGui.QMessageBox.Ok)
    buutom_ok.setText("Запустить")
    buutom_close = msg.button(QtGui.QMessageBox.Close)
    buutom_close.setText("Закрыть")
    msg.setStyleSheet(css)
    ret = msg.exec_()
    if ret == QtGui.QMessageBox.Ok:
        if get_exel_process() is False:
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
    if get_exel_process() is False:
        show_message(css, icon)
    w = AppWindow()
    w.setStyleSheet(css)
    w.resize(250, 150)
    w.move(300, 300)
    w.setWindowTitle("SartoriusScalesApp")
    w.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    create_app()