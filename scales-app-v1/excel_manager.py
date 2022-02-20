from win32com.client import dynamic, GetObject
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
        if self.app is None:
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

    def write_values(self, date, weight, date_cell, weight_cell):
        if self.current_book is None:
            logging.warning("None current book or none current sheet")
            return
        else:
            current_sheet = self.current_book.ActiveSheet
            active_cell = self.app.ActiveCell.GetAddress()
            lineNum = active_cell.split("$")[2]
            weight_current_cell = weight_cell + lineNum
            date_current_cell = date_cell + lineNum
            logging.info('writing to cells ' + active_cell + " " + weight_current_cell)
            current_sheet.Range(date_current_cell).Value = date
            current_sheet.Range(weight_current_cell).Value = weight
            self.current_book.Save()
