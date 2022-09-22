import sys
import time
from db_logger.logger import logger
from logger import log
import win32com.client
from utils.db import AppLaunchStatus
from config.base import SAP_PATH


class SapGui:
    def __init__(self, SystemName, SapLogin, SapPassword, username):
        self.__SystemName = SystemName
        self.__SapLogin = SapLogin
        self.__SapPassword = SapPassword
        self.username = username
        self.session = None

    def get_SAP_object(self):
        logs = AppLaunchStatus('get_SAP_object')
        wsh = win32com.client.Dispatch('WScript.Shell')
        wsh.Run(SAP_PATH, 1)
        while not wsh.appactivate('SAP Logon '):
            time.sleep(1)
        del wsh
        oSapGui = win32com.client.GetObject('SAPGUI')
        oApp = oSapGui.GetScriptingEngine
        oConn = oApp.OpenConnection(self.__SystemName, True)
        session = oConn.Children(0)

        session.FindById('wnd[0]/usr/txtRSYST-BNAME').text = self.__SapLogin
        session.FindById('wnd[0]/usr/pwdRSYST-BCODE').text = self.__SapPassword
        session.FindById('wnd[0]/usr/txtRSYST-LANGU').text = 'RU'
        session.FindById('wnd[0]').SendVKey(0)

        try:
            session.FindById('wnd[0]/usr/pwdRSYST-BCODE')
            self.kill_win_process('saplogon.exe')
            logs.set_end_status_work()
            sys.exit('ОШИБКА авторизации SAP!')
        except:
            session.findById('wnd[1]/tbar[0]/btn[0]').press()
            if session.activewindow.name == 'wnd[1]':
                self.kill_win_process('saplogon.exe')
                logs.set_end_status_work()
                logger.set_log('', 'SAP', 'sap_gui','getSAPobject', 'Error', 'Авторизация','Истек срок действия пароля или сессия SAP занята другим пользователем')
                log.error('Авторизация SAP GUI - Истек срок действия пароля или сессия SAP занята другим пользователем')
                sys.exit('ОШИБКА авторизации SAP! Истек срок действия пароля или сессия SAP занята другим пользователем')
            oConn = None
            oApp = None
            oSapGui = None
            logger.set_log('', 'SAP', 'sap_gui','getSAPobject', 'OK', 'Авторизация', 'Авторизация в SAP успешно завершена')
            log.info('Авторизация в SAP успешно завершена')
            del logs
            self.session = session

    def start_transaction(self, tname):
        self.session.StartTransaction(tname)

    def end_transaction(self):
        self.session.EndTransaction()

    def exit(self):
        self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    def save_export(self, fPath, fName, killexcel: bool):
        try:
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = fPath
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fName

            if killexcel:
                # на всякий случай грохаем эксель, если есть
                self.kill_win_process('excel.exe')

            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            if killexcel:
                time.sleep(3)

                self.kill_win_process('excel.exe')
        except:
            logs = AppLaunchStatus('Payments processing')
            self.kill_win_process('saplogon.exe')
            logs.set_end_status_work()
            logger.set_log('', 'SAP', 'sap_gui','save_export', 'Error', 'Сохранение выгрузки','ОШИБКА сохранения выгрузки!')
            log.error('Сохранение выгрузки - ОШИБКА сохранения выгрузки!')
            sys.exit('ОШИБКА сохранения выгрузки из SAP!')


    def kill_win_process(self, process_name):
        winServ = win32com.client.GetObject(r'winmgmts:\\.\root\CIMV2')
        props = winServ.ExecQuery("select * from Win32_Process where name = '" + process_name + "'")
        for proc in props:
            parm = proc.ExecMethod_('GetOwner')
            user = parm.Properties_('User').Value
            if self.username == user:
                proc.Terminate
