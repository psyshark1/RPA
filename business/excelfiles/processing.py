import os
import re
from time import time
from tracemalloc import start

import win32com.client

from db_logger.logger import logger
from logger import log
from config.base import EXCEL_SQL_CONNECTION_STRING, TEMP_DIR,  PREFERDATE_REPORT_NAME

connSQL = win32com.client.Dispatch('ADODB.Connection')
connSQL.CommandTimeout = 600
connSQL.ConnectionTimeout = 300

def PDOcheking(fifm_rep, exc_pay, session):
    '''Выполняет проверку на возможность создания PDO'''
    exc_pay_data = get_exc_pay_data(exc_pay, session)
    if len(exc_pay_data) != 0:
        set_comment_fifm_contract_exc(exc_pay_data, fifm_rep, session)
        calculate_comments(fifm_rep,'Запись в исключениях', session)
    del exc_pay_data

    set_comment_fifm_summ_less_1(fifm_rep)
    calculate_comments(fifm_rep,'Сумма < 1', session)

    set_comment_fifm_not_init212(fifm_rep)
    calculate_comments(fifm_rep, 'Основной счет не INIT212', session)

    set_comment_fifm_last_day_of_month(fifm_rep, session)

    set_comment_fifm_is_PDO(fifm_rep, session)

def Getting_data_for_PDO(fifm_rep, alter):

    contracts = get_fm_contracts_with_reqs(fifm_rep)
    if len(contracts) != 0:
        contracts = get_alternatives(alter,contracts)
        set_alternatives(fifm_rep, contracts)
    return get_prepared_data(fifm_rep)

def Update_pdo(fifm_rep, data):
    '''обновляет статусы созданных карточек PDO'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]
    connSQL.Open()

    for dat in data:
        if dat.get(fifm_rep.pdo_number):
            qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = '{dat.get(fifm_rep.comment)}', [{fifm_rep.pdo_number}] = '{dat.get(fifm_rep.pdo_number)}', " \
            f"[{fifm_rep.pdo_status}] = '{dat.get(fifm_rep.pdo_status)}' WHERE [{fifm_rep.fm_contract}] = '{str(dat.get(fifm_rep.fm_contract))}' AND [{fifm_rep.doc_position}] = '{dat.get(fifm_rep.doc_position)}'"
            connSQL.Execute(qry)

    connSQL.Close()

def Update_prefData(fifm_rep, data):
    '''проставляет предпочтительную лату оплаты созданных карточек PDO'''
    content = __read_rep_prefDate(str(TEMP_DIR) + '\\' + PREFERDATE_REPORT_NAME)

    for dat in data:
        if dat.get(fifm_rep.pdo_number):
            result = re.findall(r'\|' + dat.get(fifm_rep.pdo_number) + r'\|[\s\S]+?\|(\d{2}\.\d{2}\.\d{4})\|', content, flags=re.IGNORECASE)
            if len(result) == 0:
                logger.set_log(dat.get(fifm_rep.fm_contract), 'Excel', 'processing', 'Update_prefData', 'Warning', 'Получение срока оплаты', f'Срок оплаты по карточке {str(dat.get(fifm_rep.fm_contract))} не найден')
                log.warning(f'{dat.get(fifm_rep.fm_contract)} - Получение срока оплаты - Срок оплаты по карточке {str(dat.get(fifm_rep.fm_contract))} не найден')
                continue
            dat.update({fifm_rep.preliminary_pay_date: result[0]})
            set_preliminary_pay_date(fifm_rep, dat.get(fifm_rep.fm_contract), dat.get(fifm_rep.preliminary_pay_date))
            logger.set_log(dat.get(fifm_rep.fm_contract), 'Excel', 'processing', 'Update_prefData', 'OK', 'Получение срока оплаты', f'Срок оплаты по карточке {str(dat.get(fifm_rep.fm_contract))} получен')
            log.info(f'{dat.get(fifm_rep.fm_contract)} - Получение срока оплаты - Срок оплаты по карточке {str(dat.get(fifm_rep.fm_contract))} получен')


def get_exc_pay_data(exc_pay, session) -> tuple:

    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + exc_pay.fPath + EXCEL_SQL_CONNECTION_STRING[1]
    qry = f"SELECT DISTINCT [{exc_pay.contract}] FROM [{exc_pay.tblName}$] WHERE [{exc_pay.contract}] NOT LIKE '% %'"

    data = []
    connSQL.Open()

    rst = connSQL.Execute(qry)
    try:
        rst[0].movefirst
    except:
        pass

    start_time = time()
    while not rst[0].EOF:
        data.append(int(rst[0].Fields(exc_pay.contract).value))
        rst[0].movenext

        if time() - start_time >= 600:
            session.FindById("wnd[0]").sendVKey(0)
            start_time = time()

    connSQL.Close()
    return tuple(data)

def set_comment_fifm_contract_exc(exc_pay_data, fifm_rep, session):
    '''проставляет комментарий с договорами в исключениях'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]
    connSQL.Open()

    start_time = time()
    for cred in exc_pay_data:
        qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = 'Запись в исключениях' WHERE [{fifm_rep.fm_contract}] = '{str(cred)}'"
        connSQL.Execute(qry)

        if time() - start_time >= 600:
            session.FindById("wnd[0]").sendVKey(0)
            start_time = time()

    connSQL.Close()

def set_comment_fifm_summ_less_1(fifm_rep):
    '''проставляет комментарий с суммами < 1'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = IIF([{fifm_rep.comment}] IS NULL,'Сумма < 1', [{fifm_rep.comment}] + ' Сумма < 1') WHERE [{fifm_rep.fm_summ}] <= 1"

    connSQL.Open()
    connSQL.Execute(qry)
    connSQL.Close()

def set_comment_fifm_not_init212(fifm_rep):
    '''проставляет комментарий с основной счет не 212'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = IIF([{fifm_rep.comment}] IS NULL,'Основной счет не INIT212',[{fifm_rep.comment}] + ' Основной счет не INIT212') WHERE [{fifm_rep.main_acc3223}] <> 'INIT212'"

    connSQL.Open()
    connSQL.Execute(qry)
    connSQL.Close()

def set_comment_fifm_last_day_of_month(fifm_rep, session):
    '''проставляет комментарий последний день месяца'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = IIF([{fifm_rep.comment}] IS NULL,'Контрольная дата равна последнему дню месяца',[{fifm_rep.comment}] +' Контрольная дата равна последнему дню месяца') " \
    f"WHERE [{fifm_rep.pay_date}] LIKE '31.01%' OR " \
    f"[{fifm_rep.pay_date}] LIKE '28.02%' OR [{fifm_rep.pay_date}] LIKE '29.02%' OR [{fifm_rep.pay_date}] LIKE '31.03%' OR " \
    f"[{fifm_rep.pay_date}] LIKE '30.04%' OR [{fifm_rep.pay_date}] LIKE '31.05%' OR [{fifm_rep.pay_date}] LIKE '30.06%' OR " \
    f"[{fifm_rep.pay_date}] LIKE '31.07%' OR [{fifm_rep.pay_date}] LIKE '31.08%' OR [{fifm_rep.pay_date}] LIKE '30.09%' OR " \
    f"[{fifm_rep.pay_date}] LIKE '31.10%' OR [{fifm_rep.pay_date}] LIKE '30.11%' OR [{fifm_rep.pay_date}] LIKE '31.12%'"

    connSQL.Open()
    connSQL.Execute(qry)
    connSQL.Close()

    qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}] AS cont FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.comment}] LIKE '%Контрольная%'"

    connSQL.Open()
    rst = connSQL.Execute(qry)
    try:
        rst[0].movefirst
    except:
        pass

    start_time = time()
    while not rst[0].EOF:
        logger.set_log(f'{rst[0].Fields("cont").value}', 'Excel', 'processing','set_comment_fifm_last_day_of_month', 'Warning', 'Проверка на возможность создания ПДО', 'Контрольная дата равна последнему дню месяца')
        log.warning(f'{rst[0].Fields("cont").value} - Проверка на возможность создания ПДО - Контрольная дата равна последнему дню месяца')
        rst[0].movenext

        if time() - start_time >= 600:
            session.FindById("wnd[0]").sendVKey(0)
            start_time = time()

    connSQL.Close()

def set_comment_fifm_is_PDO(fifm_rep, session):
    '''проставляет комментарий есть номер PDO'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = IIF([{fifm_rep.comment}] IS NULL,'Есть номер PDO',[{fifm_rep.comment}] + ' Есть номер PDO') WHERE [{fifm_rep.pdo_number}] IS NOT NULL"

    connSQL.Open()
    connSQL.Execute(qry)
    connSQL.Close()

    connSQL.Open()

    qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}] AS cont FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.comment}] LIKE '%Есть номер PDO%'"
    rst = connSQL.Execute(qry)
    try:
        rst[0].movefirst
    except:
        pass

    start_time = time()
    while not rst[0].EOF:
        logger.set_log(f'{rst[0].Fields("cont").value}', 'Excel', 'processing', 'set_comment_fifm_is_PDO', 'Warning', 'Проверка на возможность создания ПДО', 'Есть номер PDO')
        log.warning(f'{rst[0].Fields("cont").value} - Проверка на возможность создания ПДО - Есть номер PDO')
        rst[0].movenext

        if time() - start_time >= 600:
            session.FindById("wnd[0]").sendVKey(0)
            start_time = time()

    connSQL.Close()

def set_comment(fifm_rep, contract, text, position=None, pdo_number=None, pdo_status=None):
    '''проставляет комментарий'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    if position is None:
        qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = IIF([{fifm_rep.comment}] IS NULL,'{text}',[{fifm_rep.comment}] + ' {text}') WHERE [{fifm_rep.fm_contract}] = '{str(contract)}'"
    else:
        if pdo_number is not None and pdo_status is not None:
            qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = IIF([{fifm_rep.comment}] IS NULL,'{text}',[{fifm_rep.comment}] + ' {text}'), [{fifm_rep.pdo_number}] = '{pdo_number}', [{fifm_rep.pdo_status}] = '{pdo_status}' " \
                  f"WHERE [{fifm_rep.fm_contract}] = '{str(contract)}' AND [{fifm_rep.doc_position}] = '{str(position)}'"
        else:
            qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.comment}] = IIF([{fifm_rep.comment}] IS NULL,'{text}',[{fifm_rep.comment}] + ' {text}') WHERE [{fifm_rep.fm_contract}] = '{str(contract)}' AND [{fifm_rep.doc_position}] = '{str(position)}'"

    connSQL.Open()
    connSQL.Execute(qry)
    connSQL.Close()

def set_pay_requisites(fifm_rep, contract, rqs_code):
    '''проставляет платежный реквизит'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.requisites_code}] = '{rqs_code}' WHERE [{fifm_rep.fm_contract}] = '{str(contract)}' AND [{fifm_rep.comment}] IS NULL"

    connSQL.Open()
    connSQL.Execute(qry)
    connSQL.Close()

def set_preliminary_pay_date(fifm_rep, contract, preliminary_pay_date):
    '''проставляет предпочтительную дату оплаты'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.preliminary_pay_date}] = '{preliminary_pay_date}' WHERE [{fifm_rep.fm_contract}] = '{str(contract)}'"

    connSQL.Open()
    connSQL.Execute(qry)
    connSQL.Close()

def get_fm_contracts(fifm_rep):
    '''получает договоры '''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}], [{fifm_rep.control_date}], [{fifm_rep.doc_position}] FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.comment}] IS NULL"

    contracts = []
    connSQL.Open()
    rst = connSQL.Execute(qry)
    try:
        rst[0].movefirst
    except:
        pass

    while not rst[0].EOF:
        tmp = {}
        for field in rst[0].Fields:
            tmp[field.Name] = field.value

        contracts.append(tmp)

        rst[0].movenext

    connSQL.Close()
    return  contracts

def get_fm_contracts_with_reqs(fifm_rep):
    '''получает договоры с реквизитами'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}] FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.requisites_code}] IS NOT NULL"
    contracts = []

    connSQL.Open()
    rst = connSQL.Execute(qry)
    try:
        rst[0].movefirst
    except:
        pass

    while not rst[0].EOF:
        contracts.append(int(rst[0].Fields(fifm_rep.fm_contract).value))
        rst[0].movenext

    connSQL.Close()
    return  tuple(contracts)

def calculate_comments(fifm_rep, text_comment, session):
    '''определяет договоры с комментарием для записи в лог'''
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    if text_comment == 'Запись в исключениях':
        qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}] AS cont FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.comment}] LIKE '%{text_comment}%'"
    elif text_comment == 'Основной счет не INIT212':
        qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}] AS cont, [{fifm_rep.doc_position}] AS pos FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.comment}] LIKE '%{text_comment}%'"
    else:
        qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}] AS cont FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.comment}] LIKE '%{text_comment}%'"

    connSQL.Open()
    rst = connSQL.Execute(qry)
    try:
        rst[0].movefirst
    except:
        pass

    start_time = time()
    while not rst[0].EOF:
        if text_comment == 'Запись в исключениях':
            logger.set_log(f'{rst[0].Fields("cont").value}', 'Excel', 'processing', 'calculate_comments', 'Warning', 'Проверка на возможность создания ПДО',f'{text_comment}')
            log.warning(f'{rst[0].Fields("cont").value} - Проверка на возможность создания ПДО - {text_comment}')
        elif text_comment == 'Основной счет не INIT212':
            logger.set_log(f'{rst[0].Fields("cont").value}', 'Excel', 'processing', 'calculate_comments', 'Warning', 'Проверка на возможность создания ПДО',f'{text_comment} с позицией {rst[0].Fields("pos").value}')
            log.warning(f'{rst[0].Fields("cont").value} - Проверка на возможность создания ПДО - {text_comment} с позицией {rst[0].Fields("pos").value}')
        else:
            logger.set_log(f'{rst[0].Fields("cont").value}', 'Excel', 'processing', 'calculate_comments', 'Warning', 'Проверка на возможность создания ПДО',f'{text_comment}')
            log.warning(f'{rst[0].Fields("cont").value} - Проверка на возможность создания ПДО - {text_comment}')
        rst[0].movenext

        if time() - start_time >= 600:
            session.FindById("wnd[0]").sendVKey(0)
            start_time = time()
        #break

    connSQL.Close()

def get_alternatives(alter, contracts):
    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + alter.fPath + EXCEL_SQL_CONNECTION_STRING[1]
    connSQL.Open()

    alternatives = {}

    for contract in contracts:

        qry = f"SELECT DISTINCT [{alter.contract}], [{alter.creditor}], [{alter.creditor_code}], [{alter.requisites_code}] FROM [{alter.tblName}$] WHERE [{alter.contract}] = {contract} " \
        f"AND [{alter.requisites_code}] IS NOT NULL AND [{alter.creditor_code}] IS NOT NULL AND ISNUMERIC([{alter.creditor_code}])"

        rst = connSQL.Execute(qry)
        try:
            rst[0].movefirst
        except:
            pass

        while not rst[0].EOF:
            tmp = {}

            tmp[int(rst[0].Fields(alter.creditor).value)] = (int(rst[0].Fields(alter.creditor_code).value), int(rst[0].Fields(alter.requisites_code).value))

            alternatives[contract] = tmp
            rst[0].movenext

    connSQL.Close()
    return alternatives

def set_alternatives(fifm_rep, contracts):

    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]
    connSQL.Open()

    for key, value in contracts.items():
        for k, v in value.items():
            qry = f"UPDATE [{fifm_rep.tblName}$] SET [{fifm_rep.alt_recipient}] = '{str(v[0])}', [{fifm_rep.alt_requisites_code}] = '{v[1]}' WHERE " \
            f"[{fifm_rep.fm_contract}] = '{str(key)}' AND [{fifm_rep.creditor}] = '{str(k)}'"
            connSQL.Execute(qry)

    connSQL.Close()

def get_prepared_data(fifm_rep):

    connSQL.ConnectionString = EXCEL_SQL_CONNECTION_STRING[0] + fifm_rep.fPath + EXCEL_SQL_CONNECTION_STRING[1]

    qry = f"SELECT DISTINCT [{fifm_rep.fm_contract}],[{fifm_rep.fm_be}],[{fifm_rep.creditor}],[{fifm_rep.control_date}], " \
    f"[{fifm_rep.requisites_code}],[{fifm_rep.doc_position}],[{fifm_rep.alt_recipient}],[{fifm_rep.alt_requisites_code}] " \
    f"FROM [{fifm_rep.tblName}$] WHERE [{fifm_rep.requisites_code}] IS NOT NULL"

    data = []

    connSQL.Open()
    rst = connSQL.Execute(qry)
    try:
        rst[0].movefirst
    except:
        pass

    while not rst[0].EOF:
        tmp = {}
        for field in rst[0].Fields:
            tmp[field.Name] = field.value

        data.append(tmp)

        rst[0].movenext

    connSQL.Close()

    return data

def __read_rep_prefDate(fpath):
    f = open(fpath, 'r')
    cnt = f.read()
    f.close()
    os.remove(fpath)
    return cnt
