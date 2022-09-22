import os
import shutil
import winreg
from time import sleep
#from turtle import back

import win32com.client
from pywinauto.application import Application

from logger import log

class Register:
    def __init__(self):
        self.oxl = win32com.client.Dispatch('excel.application')
        self.oxl.visible = False
        self.oxl.DisplayAlerts = False

    def close(self, save: bool):
        self.xl.Close(save)

    def quit(self):
        self.oxl.Quit()
        self.oxl = None

    @staticmethod
    def open_on_error(method):
        def wrapper(obj_self, *args, **kwargs):
            try:
                method(obj_self, *args, **kwargs)
            except BaseException as e:
                if -2146827284 in e.args[2] or -2147352567 in e.args[2]:
                    xl = Application(backend='uia').start(f'C:\\Program Files\\Microsoft Office\\Office15\\EXCEL.exe {obj_self.fPath}')
                    xl.Dialog2.Button1.invoke()
                    sleep(2)
                    xl.kill()
                    method(obj_self, *args, **kwargs)
                else:
                    log.exception(e)
        return wrapper


class FiFMReport(Register):
    '''Отчет план-факт по затратам FI-FM'''
    def __init__(self, fPath):
        super().__init__()
        self.xl = None
        self.fPath = fPath
        self.tblName = None
        self.fm_be = 'FM БЕ'
        self.fm_contract = 'FM Номер документа'
        self.fm_summ = 'FM Сумма (RUB)'
        self.creditor = 'Кредитор'
        self.control_date = 'Контрольная дата'
        self.doc_position = 'Позиция документа'
        self.pay_date = 'Срок оплаты'
        self.main_acc3223 = 'Основной счет (32/23)'
        self.pdo_number = 'Номер PDO'
        self.pdo_status = 'Статус PDO'
        self.comment = 'Комментарий'
        self.requisites_code = 'Код реквизитов'
        self.alt_recipient = 'Альтернативный получатель'
        self.alt_requisites_code = 'Альтернативный код реквизитов'
        self.preliminary_pay_date = 'Пред срок оплаты'

    def __get_tblName(self):
        self.tblName = self.xl.sheets(1).Name

    def add_columns(self):
        #self.xl.sheets(1).cells(1, 10).value = self.main_acc3223
        self.xl.sheets(1).cells(1, 14).value = self.comment
        self.xl.sheets(1).cells(1, 15).value = self.requisites_code
        self.xl.sheets(1).cells(1, 16).value = self.alt_recipient
        self.xl.sheets(1).cells(1, 17).value = self.alt_requisites_code
        self.xl.sheets(1).cells(1, 18).value = self.preliminary_pay_date
        self.__get_tblName()

    def rename(self, new_name):
        os.rename(self.fPath, new_name)
        self.fPath = new_name

    def move(self, src, dst):
        shutil.move(src, dst)
        self.fPath = dst

    @Register.open_on_error
    def open(self):
        strPath = "SOFTWARE\\Microsoft\\Office\\15.0\\Excel\\Resiliency\\DisabledItems\\"
        registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, strPath, 0, winreg.KEY_READ)
        cnt = winreg.QueryInfoKey(registry_key)[1]
        if cnt > 0:
            value = winreg.EnumValue(registry_key, 0)[0]
            registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, strPath, 0, winreg.KEY_WRITE)
            winreg.DeleteValue(registry_key, value)
        self.xl = self.oxl.Workbooks.Open(self.fPath, 0, 0)

    def saveas(self):
        self.fPath = self.fPath[:-5] + 'xlsx'
        self.xl.SaveAs(self.fPath, 51, None, None, None, False)

class ExceptionPayment(Register):
    '''Исключения для оплаты'''
    def __init__(self, fPath):
        super().__init__()
        self.xl = None
        self.fPath = fPath
        self.tblName = None
        self.contract = 'Договор'
        self.creditor = 'Кредитор'

    def __get_tblName(self):
        self.tblName = self.xl.sheets(1).Name

    @Register.open_on_error
    def open(self):
        self.xl = self.oxl.Workbooks.Open(self.fPath, 0, 1)
        self.__get_tblName()

    def copy(self, src, dst, fname):
        shutil.copy(src, dst)
        self.fPath = dst + '\\' + fname

class AlternateRecipient(Register):
    '''Альтернативные получатели'''
    def __init__(self, fPath):
        super().__init__()
        self.xl = None
        self.fPath = fPath
        self.tblName = None
        self.contract = 'Договор'
        self.creditor = 'Кредитор'
        self.creditor_code = 'Код кредитора'
        self.requisites_code = 'Код реквизитов'

    def __get_tblName(self):
        self.tblName = self.xl.sheets(1).Name

    '''def rename_column(self):
        self.xl.sheets(1).cells(1, 7).value = self.requisites_code
        self.__get_tblName()'''

    @Register.open_on_error
    def open(self):
        self.xl = self.oxl.Workbooks.Open(self.fPath, 0, 1)
        self.__get_tblName()

    def copy(self, src, dst, fname):
        shutil.copy(src, dst)
        self.fPath = dst + '\\' + fname

class ApproveBindPositions(Register):
    '''эксель-представление позиций для связывания'''
    def __init__(self, fPath):
        super().__init__()
        self.xl = None
        self.fPath = fPath

    @Register.open_on_error
    def open(self):
        self.xl = self.oxl.Workbooks.Open(self.fPath, 0, 1)

    def delete_row_processing(self, doc_num,pos_num):

        i=2;r=0;chk = False
        while self.xl.sheets(1).cells(i, 1).value is not None:
            if self.xl.sheets(1).cells(i, 1).value == doc_num and self.xl.sheets(1).cells(i, 2).value == pos_num:
                r = i - 2
                chk = True
            i+=1
        return (r, i-2, chk)
