import os

import win32com.client

from business.sap_business import SapBusiness
from business.excelfiles.registers import FiFMReport, ExceptionPayment, AlternateRecipient
from config.base import AUTH_DATA, TEMP_DIR, WORK_DIR, FIFM_REPORT_NAME, ALTER_REPORT_NAME, EXCEPTION_PAY_NAME, RESULT_DIR, RESULT_REPORT_NAME, LOG_DIR
from business.excelfiles.processing import PDOcheking, Getting_data_for_PDO, Update_pdo
from db_logger.logger import logger
from logger import log
from utils.db.app_launch_status import AppLaunchStatus

def do_payments_processing():

    logger.set_log('', 'Python', 'main', 'do_payments_processing', 'Info', 'Исполнение', 'Начало работы')
    log.add(level='INFO', dir_path=LOG_DIR)
    log.info('Исполнение - Начало работы')
    ls = AppLaunchStatus('Payments processing')
    ls.set_start_status_work()

    cur_user = win32com.client.GetObject('LDAP://' + win32com.client.Dispatch('adsysteminfo').username).samaccountname.lower()
    sapgui = SapBusiness('101. EMP - Продуктив', AUTH_DATA['sap_gui_login'], AUTH_DATA['sap_gui_pass'], cur_user)
    sapgui.get_report_fifm('ZAG_COMPCOST')
    fifm_rep = FiFMReport(str(TEMP_DIR) + '\\' + FIFM_REPORT_NAME)
    fifm_rep.open()
    fifm_rep.add_columns()
    fifm_rep.saveas()
    fifm_rep.close(False)#должно быть False
    fifm_rep.quit()
    os.remove(str(TEMP_DIR) + '\\' + FIFM_REPORT_NAME)#удаляет mhtml

    if os.path.isfile(str(TEMP_DIR) + '\\' + EXCEPTION_PAY_NAME):
        os.remove(str(TEMP_DIR) + '\\' + EXCEPTION_PAY_NAME)
    exc_pay = ExceptionPayment(WORK_DIR + '\\' + EXCEPTION_PAY_NAME)
    exc_pay.copy(WORK_DIR + '\\' + EXCEPTION_PAY_NAME, str(TEMP_DIR), EXCEPTION_PAY_NAME)
    exc_pay.open()
    exc_pay.close(False)
    exc_pay.quit()

    PDOcheking(fifm_rep, exc_pay, sapgui.session)
    del exc_pay

    sapgui.get_bank_details('ZDR3',fifm_rep)

    if os.path.isfile(str(TEMP_DIR) + '\\' + ALTER_REPORT_NAME):
        os.remove(str(TEMP_DIR) + '\\' + ALTER_REPORT_NAME)
    alter = AlternateRecipient(WORK_DIR + '\\' + ALTER_REPORT_NAME)
    alter.copy(WORK_DIR + '\\' + ALTER_REPORT_NAME, str(TEMP_DIR), ALTER_REPORT_NAME)
    alter.open()
    #alter.rename_column()
    alter.close(False)#должно быть False, теперь только чтение имени листа
    alter.quit()

    data = Getting_data_for_PDO(fifm_rep,alter)

    sapgui.create_PDO('ZPDO1', data, fifm_rep)

    Update_pdo(fifm_rep, data)

    sapgui.get_prefDate('Z012_REPAPPDATA', data, fifm_rep)

    fifm_rep.rename(str(TEMP_DIR) + '\\' + RESULT_REPORT_NAME)
    try:
        fifm_rep.move(str(TEMP_DIR) + '\\' + RESULT_REPORT_NAME, RESULT_DIR + '\\' + RESULT_REPORT_NAME)
    except:
        logger.set_log('', 'Python', 'main', 'do_payments_processing', 'Error', 'Исполнение',f'Ошибка перемещения итогового отчета в {RESULT_DIR}')
        log.error(f'Исполнение - Ошибка перемещения итогового отчета в {RESULT_DIR}')
        pass

    sapgui.exit()
    sapgui.kill_win_process('saplogon.exe')

    ls.set_cnt_request(len(data))
    ls.update_cnt_good(len(data))
    ls.set_end_status_work()

    logger.set_log('', 'Python', 'main', 'do_payments_processing', 'Info', 'Исполнение', 'Штатное завершение работы')
    log.info('Исполнение - Штатное завершение работы')
