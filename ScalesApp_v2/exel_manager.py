from win32com.client import dynamic
import logging


class ExcelManagerMeta(type):
    def __init__(self, name, bases, dic):
        self.__instance = None
        super().__init__(name, bases, dic)

    def __call__(cls, *args, **kwargs):
        if cls.__instance:
            return cls.__instance
        obj = cls.__new__(cls)
        obj.__init__(*args, **kwargs)
        cls.__instance = obj
        return obj


class ExcelManager(metaclass=ExcelManagerMeta):
    def __init__(self):
        logging.info('Create excel manager')
        self.app = None
        self.current_book = None
        self.active_books = self.get_books()
        self.name_active_book = ""

    def get_books(self):
        logging.info('Loading active books')
        active_books = []
        if self.app == None:
            active_books = [workbook.Name for workbook in dynamic.Dispatch("Excel.Application").Workbooks]
        else:
            active_books = self.active_books
        if len(active_books) == 0:
            logging.error('No active books')
        return active_books

    def update_books(self):
        logging.info('Update active books')
        active_books = [workbook.Name for workbook in dynamic.Dispatch("Excel.Application").Workbooks]
        if len(active_books) == 0:
            logging.error('No active books')
        else:
            self.active_books = active_books
        return active_books

    def set_active_book(self, name_book):
        if name_book not in self.active_books:
            return
        logging.info('Set active book')
        self.app = dynamic.Dispatch("Excel.Application")
        self.current_book = self.app.Workbooks(name_book)
        self.name_active_book = name_book

    def clear_parametrs(self):
        self.current_sheet = None
        self.letter = None
        self.num_line = None
        logging.info("Clear excel parametrs")

    def set_parametr_write(self):
        if self.current_book == None:
            logging.warning("None current book")
            return
        else:
            self.current_sheet = self.current_book.ActiveSheet
            active_cell = self.app.ActiveCell.GetAddress()
            self.letter = active_cell.split("$")[1]
            self.num_line = active_cell.split("$")[2]
            logging.info("Set active cell " + active_cell)

    def write_values(self, weight, count):
        if self.current_book == None or self.current_sheet == None:
            logging.warning("None current book or none current sheet")
            return
        else:
            self.current_sheet.Range(self.letter + str(int(self.num_line) + count)).Value = weight
            self.current_book.Save()
