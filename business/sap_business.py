import os
import sys
from datetime import datetime, timedelta
from time import sleep
import calendar

from business.sap.sap_gui import SapGui
from business.excelfiles.registers import ApproveBindPositions
from business.excelfiles.processing import get_fm_contracts, set_comment, set_pay_requisites, Update_prefData
from config.base import TEMP_DIR, APPROVE_REPORT_NAME, FIFM_REPORT_NAME, PREFERDATE_REPORT_NAME, SAP_REPORT_FORMAT_ROW
from utils.db import AppLaunchStatus
from db_logger.logger import logger
from logger import log

class SapBusiness(SapGui):
    def __init__(self, SystemName, SapLogin, SapPassword, username):
        super().__init__(SystemName, SapLogin, SapPassword, username)
        super().get_SAP_object()

    def get_report_fifm(self, tname):

        super().start_transaction(tname)

        self.session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = 'W*'
        self.session.findById("wnd[0]/usr/ctxtS_BELNR-LOW").text = '32*'
        self.session.findById("wnd[0]/usr/ctxtS_FKBER-LOW").text = '421630426'
        self.session.findById("wnd[0]/usr/btn%_S_FDATK_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").select()
        self.session.findById("wnd[0]/usr/btn%_S_FDATK_%_APP_%-VALU_PUSH").press()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL-SOP_I[0,0]").press()
        self.session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = 2
        self.session.findById("wnd[2]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1,0]").text = (datetime.today() + timedelta(days=20)).strftime('%d.%m.%Y')
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/radERLKZ_2").select()
        self.session.findById("wnd[0]/usr/ctxtS_LNRZA-LOW").text = '01.12.2021'
        new_date = self.__add_months(self.__last_day_of_month(datetime.today().date()), 2, False)
        self.session.findById("wnd[0]/usr/ctxtS_LNRZA-HIGH").text = new_date.strftime('%d.%m.%Y')
        del new_date
        self.session.findById("wnd[0]/usr/ctxtP_VARALV").setFocus()
        #self.session.findById("wnd[0]/usr/ctxtP_VARALV").text = tname #"ZAG_COMPCOST"
        self.session.findById('wnd[0]').sendVKey(4)
        try:
            self.session.findById(f'wnd[1]/usr/lbl[1,{SAP_REPORT_FORMAT_ROW}]').setFocus()
        except:
            super().kill_win_process('saplogon.exe')
            logs = AppLaunchStatus('get_report_fifm')
            logs.set_end_status_work()
            logger.set_log('', 'SAP', 'sap_business', 'get_report_fifm', 'Error', 'Получение отчета Сравнения план-факт по затратам FI-FM','Ошибка выбора формата отчета')
            log.error('Получение отчета Сравнения план-факт по затратам FI-FM - Ошибка выбора формата отчета')
            sys.exit('Ошибка выбора формата отчета')
        self.session.findById('wnd[1]').sendVKey(2)
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

        self.session.findById("wnd[0]/tbar[1]/btn[43]").press()

        super().save_export(str(TEMP_DIR), FIFM_REPORT_NAME, True)

        super().end_transaction()
        logger.set_log('', 'SAP', 'sap_business', 'get_report_fifm', 'OK', 'Получение отчета Сравнения план-факт по затратам FI-FM', 'Отчет успешно сохранен')
        log.info('Получение отчета Сравнения план-факт по затратам FI-FM - Отчет успешно сохранен')

    def get_bank_details(self, tname, fifm_rep):
        '''получает банковские реквизиты'''

        contracts = get_fm_contracts(fifm_rep)
        for contract in contracts:
            if datetime.now().hour == 3 and datetime.now().minute >= 0 <= 19:
                sleep(60*20)

            logger.set_log(contract.get(fifm_rep.fm_contract), 'SAP', 'sap_business','get_bank_details', 'Info', 'Получение банковских реквизитов','')
            log.info(f'{contract.get(fifm_rep.fm_contract)} - Получение банковских реквизитов')
            chk = False

            for month_depth in range(-1,-3,-1):
                if chk: break
                if month_depth == -1 and not chk:
                    super().start_transaction(tname)

                    self.session.findById("wnd[0]/usr/ctxtDRAW-DOKNR").text = contract.get(fifm_rep.fm_contract)
                    self.session.findById('wnd[0]').sendVKey(0)
                    self.session.findById('wnd[0]/tbar[1]/btn[5]').press()

                rslt = self.__check_control_date(contract.get(fifm_rep.control_date), month_depth)

                if not rslt[0]:
                    if month_depth == -1:
                        self.session.findById('wnd[0]/usr/tblSAPLFMFRG_TC_POSITIONS').verticalScrollbar.position = 0
                        continue
                    else:
                        set_comment(fifm_rep, contract.get(fifm_rep.fm_contract), 'Не определены реквизиты платежа', contract.get(fifm_rep.doc_position))
                        logger.set_log(contract.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'get_bank_details', 'Warning', 'Получение банковских реквизитов', f'Не определены реквизиты платежа по позиции {contract.get(fifm_rep.doc_position)}')
                        log.warning(f'{contract.get(fifm_rep.fm_contract)} - Получение банковских реквизитов - Не определены реквизиты платежа по позиции {contract.get(fifm_rep.doc_position)}')
                        self.session.findById('wnd[0]/tbar[0]/btn[15]').press()
                        self.session.findById('wnd[0]/tbar[0]/btn[15]').press()
                        break

                self.session.findById(f'wnd[0]/usr/tblSAPLFMFRG_TC_POSITIONS/txtKBLD-WTGES[1,{str(rslt[1])}]').setFocus()
                self.session.findById('wnd[0]').sendVKey(2)
                self.session.findById('wnd[0]/tbar[1]/btn[39]').press()

                t = 6; pdo_break = False
                try:
                    while True:
                        if self.session.findById(f'wnd[0]/usr/lbl[1,{str(t)}]').text == "PDO":
                            self.session.findById(f'wnd[0]/usr/lbl[13,{str(t)}]').SetFocus()
                            self.session.findById('wnd[0]').sendVKey(2)
                            self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05').select()
                            if self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/ctxtZDF_CARD_012008-BVTYP').text == '':
                                set_comment(fifm_rep, contract.get(fifm_rep.fm_contract), f'Реквизиты платежа отсутствуют на контрольной дате в {str(month_depth)} месяц(а)', contract.get(fifm_rep.doc_position))
                                self.session.findById('wnd[0]/tbar[0]/btn[3]').press()
                                logger.set_log(contract.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'get_bank_details', 'Warning', 'Получение банковских реквизитов', f'Реквизиты платежа отсутствуют на контрольной дате в {str(month_depth)} месяц(а) по позиции {contract.get(fifm_rep.doc_position)}')
                                log.warning(f'{contract.get(fifm_rep.fm_contract)} - Получение банковских реквизитов - Реквизиты платежа отсутствуют на контрольной дате в {str(month_depth)} месяц(а) по позиции {contract.get(fifm_rep.doc_position)}')
                                if month_depth == -2:
                                    pdo_break = True
                                    break
                            else:
                                set_pay_requisites(fifm_rep, contract.get(fifm_rep.fm_contract), self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/ctxtZDF_CARD_012008-BVTYP').text)
                                logger.set_log(contract.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'get_bank_details', 'OK', 'Получение банковских реквизитов',f'Найден платежный реквизит по позиции {contract.get(fifm_rep.doc_position)}')
                                log.info(f'{contract.get(fifm_rep.fm_contract)} - Получение банковских реквизитов - Найден платежный реквизит по позиции {contract.get(fifm_rep.doc_position)}')
                                self.session.findById('wnd[0]/tbar[0]/btn[3]').press()
                                chk = True
                                #wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/ctxtZDF_CARD_012008-BVTYP_EMPFK
                        t += 1
                except:
                    if month_depth == -2 and not chk:
                        logger.set_log(contract.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'get_bank_details', 'Warning','Получение банковских реквизитов', f'Не найден присвоенный документ PDO по позиции {contract.get(fifm_rep.doc_position)}')
                        log.warning(f'{contract.get(fifm_rep.fm_contract)} - Получение банковских реквизитов - Не найден присвоенный документ PDO по позиции {contract.get(fifm_rep.doc_position)}')
                        set_comment(fifm_rep, contract.get(fifm_rep.fm_contract), 'Не найден присвоенный документ PDO', contract.get(fifm_rep.doc_position))
                    pass

                if not pdo_break: self.session.findById('wnd[0]/tbar[0]/btn[3]').press()
                self.session.findById('wnd[0]/tbar[0]/btn[3]').press()
                if month_depth == -1 and not chk:
                    self.session.findById('wnd[0]/usr/tblSAPLFMFRG_TC_POSITIONS').verticalScrollbar.position = 0
                    continue
                self.session.findById('wnd[0]/tbar[0]/btn[15]').press()
                self.session.findById('wnd[0]/tbar[0]/btn[15]').press()

        #super().end_transaction()

    def create_PDO(self, tname, data, fifm_rep):

        for dat in data:
            if datetime.now().hour == 1 and datetime.now().minute >= 57 <= 59:
                sleep(60*21)

            super().start_transaction(tname)
            logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'Info', 'Создание оплатной карточки ПДО', '')
            log.info(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО')

            try:

                self.session.findById('wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/txt/DFS/STR_KEYFIELDS3-DKTXT').text = f'Оплата {dat.get(fifm_rep.control_date)}'
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC01/ssubTS012008_SCA:SAPLZDF012008:0001/ctxtZDF_CARD_012008-BUKRS').text = f'{dat.get(fifm_rep.fm_be)}'
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC01/ssubTS012008_SCA:SAPLZDF012008:0001/ctxtZDF_CARD_012008-REQTYPE_PDO').text = '1'
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC01/ssubTS012008_SCA:SAPLZDF012008:0001/ctxtZDF_CARD_012008-ZUMSK').text = 'A'
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC01/ssubTS012008_SCA:SAPLZDF012008:0001/ctxtZDF_CARD_012008-LIFNR').text = f'{dat.get(fifm_rep.creditor)}'
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC01/ssubTS012008_SCA:SAPLZDF012008:0001/ctxtZDF_CARD_012008-DOKNR_ZDR').text = f'{dat.get(fifm_rep.fm_contract)}'
                self.session.findById('wnd[0]/tbar[0]/btn[11]').press()
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05').select()

                if self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/ctxtZDF_CARD_012008-BVTYP').text != dat.get(fifm_rep.requisites_code):
                    logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'Error', 'Создание оплатной карточки ПДО', f'Платежные реквизиты не совпадают по позиции {dat.get(fifm_rep.doc_position)}')
                    log.error(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО - Платежные реквизиты не совпадают по позиции {dat.get(fifm_rep.doc_position)}')
                    set_comment(fifm_rep, dat.get(fifm_rep.fm_contract), 'Платежные реквизиты не совпадают при создании оплатной карточки', dat.get(fifm_rep.doc_position))
                    super().end_transaction()
                    continue

                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/radRB_9001_12').select()
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/ctxtZDF_CARD_012008-CURRENCY').text = "RUB"
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/txtZDF_CARD_012008-PDO_TXT_1').text = "аренду"
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/radRB_9001_23').select()
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/cmbZDF_CARD_012008-PDO_PERIOD').key = f'{int(dat.get(fifm_rep.control_date)[3:5])}'
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/txtZDF_CARD_012008-PDO_YEAR').text = f'{int(dat.get(fifm_rep.control_date)[-4:])}'

                if dat.get(fifm_rep.alt_recipient) is not None and dat.get(fifm_rep.alt_requisites_code) is not None:
                    self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/ctxtZDF_CARD_012008-EMPFK').text = dat.get(fifm_rep.alt_recipient)
                    self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC05/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9001/ctxtZDF_CARD_012008-BVTYP_EMPFK').text = dat.get(fifm_rep.alt_requisites_code)

                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC06').select()
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC06/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9004/radRB_9004_2').select()
                self.session.findById('wnd[0]/tbar[0]/btn[11]').press()
                self.session.findById('wnd[0]/usr/cntlCONT_AREA/shellcont/shell').pressToolbarContextButton('&MB_EXPORT')
                self.session.findById("wnd[0]/usr/cntlCONT_AREA/shellcont/shell").selectContextMenuItem('&XXL')

                super().save_export(str(TEMP_DIR), APPROVE_REPORT_NAME, True)

                del_rows = self.calculate_delete_rows(str(dat.get(fifm_rep.fm_contract)),dat.get(fifm_rep.doc_position))

                if del_rows[2]:
                    if del_rows[0] != 0:
                        self.session.findById("wnd[0]/usr/cntlCONT_AREA/shellcont/shell").deleteRows(f'0-{del_rows[0]-1}')
                        if del_rows[0] != del_rows[1]-1:
                            self.session.findById("wnd[0]/usr/cntlCONT_AREA/shellcont/shell").deleteRows(f'1-{del_rows[1] - del_rows[0]-1}')
                    else:
                        if del_rows[1] != 0:
                            if del_rows[1] > 1:
                                self.session.findById("wnd[0]/usr/cntlCONT_AREA/shellcont/shell").deleteRows(f'1-{del_rows[1]-1}')
                            #else:
                                #self.session.findById("wnd[0]/usr/cntlCONT_AREA/shellcont/shell").deleteRows('1')
                else:
                    logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'Warning', 'Создание оплатной карточки ПДО', f'Карточка не сформирована по позиции {dat.get(fifm_rep.doc_position)}')
                    log.warning(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО - Карточка не сформирована по позиции {dat.get(fifm_rep.doc_position)}')
                    set_comment(fifm_rep, dat.get(fifm_rep.fm_contract), 'Карточка не сформирована', dat.get(fifm_rep.doc_position))
                    #self.session.findById("wnd[0]/tbar[0]/btn[15]").press()
                    #self.session.findById("wnd[0]/tbar[0]/btn[12]").press()
                    try:
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    except:
                        raise Exception('Не найдена кнопка выхода модального окна при поиске выделенных средств')
                    #continue

                self.session.findById('wnd[0]/tbar[0]/btn[11]').press()
                self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC07').select()

                if self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC07/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9010/cntlCONTAINER_9010/shellcont/shell').RowCount == 5:

                    if self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC07/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9010/cntlCONTAINER_9010/shellcont/shell').GetCellValue(0, 'BELNR').startswith('32') and \
                    self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC07/ssubTS012008_SCA:SAPLZDF012008_CUSTOMER:9010/cntlCONTAINER_9010/shellcont/shell').GetCellValue(2, 'BELNR').startswith('23'):

                        self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC12').select()
                        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                        self.session.findById('wnd[0]/usr/tabsTS012008/tabpTS_FC12/ssubTS012008_SCA:SAPLZDF012008:2002/btnTC_APPROVAL').press()

                        if self.session.findById("wnd[0]/sbar").text.find('WorkFlow') == -1:
                            logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'Error', 'Создание оплатной карточки ПДО',
                            f'Ошибка при создании ПДО {self.session.findById("wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/txt/DFS/STR_CARDHEADER-DOKNR").text} - {self.session.findById("wnd[0]/sbar").text} по позиции {dat.get(fifm_rep.doc_position)}')
                            log.error(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО - Ошибка при создании ПДО '
                            f'{self.session.findById("wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/txt/DFS/STR_CARDHEADER-DOKNR").text} - {self.session.findById("wnd[0]/sbar").text} по позиции {dat.get(fifm_rep.doc_position)}')
                            set_comment(fifm_rep, dat.get(fifm_rep.fm_contract), f'Ошибка при создании ПДО '
                            f'{self.session.findById("wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/txt/DFS/STR_CARDHEADER-DOKNR").text} - {self.session.findById("wnd[0]/sbar").text}', dat.get(fifm_rep.doc_position))
                            super().end_transaction()
                            continue

                        dat.update({fifm_rep.pdo_number: self.session.findById('wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/txt/DFS/STR_CARDHEADER-DOKNR').text})
                        dat.update({fifm_rep.pdo_status: self.session.findById('wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/ctxt/DFS/STR_KEYFIELDS3-STABK').text})
                        dat.update({fifm_rep.comment: f'Сформирована оплатная карточка {dat.get(fifm_rep.pdo_number)}'})
                        logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'OK', 'Создание оплатной карточки ПДО', f'Сформирована оплатная карточка {dat.get(fifm_rep.pdo_number)} по позиции {dat.get(fifm_rep.doc_position)}')
                        log.info(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО - Сформирована оплатная карточка {dat.get(fifm_rep.pdo_number)} по позиции {dat.get(fifm_rep.doc_position)}')
                        # self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        super().end_transaction()
                    else:
                        logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'Error', 'Создание оплатной карточки ПДО', f'Позиции для резервирования некорректны для позиции документа {dat.get(fifm_rep.doc_position)}')
                        log.error(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО - Позиции для резервирования некорректны для позиции документа {dat.get(fifm_rep.doc_position)}')
                        set_comment(fifm_rep, dat.get(fifm_rep.fm_contract),'Позиции для резервирования некорректны', dat.get(fifm_rep.doc_position))
                        super().end_transaction()
                else:
                    logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'Error', 'Создание оплатной карточки ПДО', f'Неверное количество позиций для резервирования для позиции документа {dat.get(fifm_rep.doc_position)}')
                    log.error(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО - Неверное количество позиций для резервирования для позиции документа {dat.get(fifm_rep.doc_position)}')
                    set_comment(fifm_rep, dat.get(fifm_rep.fm_contract), 'Неверное количество позиций для резервирования', dat.get(fifm_rep.doc_position),
                                self.session.findById('wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/txt/DFS/STR_CARDHEADER-DOKNR').text, self.session.findById('wnd[0]/usr/subDFSHEADER:SAPLZDF012008:0998/ctxt/DFS/STR_KEYFIELDS3-STABK').text)
                    super().end_transaction()
            except Exception as err:
                logger.set_log(dat.get(fifm_rep.fm_contract), 'SAP', 'sap_business', 'create_PDO', 'Error',
                               'Создание оплатной карточки ПДО',
                               f'Ошибка при исоздании карточки PDO по кредитору {dat.get(fifm_rep.creditor)}')
                log.error(f'{dat.get(fifm_rep.fm_contract)} - Создание оплатной карточки ПДО - Ошибка при исоздании карточки PDO по кредитору {dat.get(fifm_rep.creditor)}')
                log.exception(err)
                super().end_transaction()

    def get_prefDate(self, tname, data, fifm_rep):
        '''Получает срок оплаты'''
        super().start_transaction(tname)

        self.session.findById("wnd[0]/usr/radP_PDO").select()
        self.session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "W*"
        self.session.findById("wnd[0]/usr/btn%_S_DOKNR_%_APP_%-VALU_PUSH").press()
        sr=0
        for dat in data:
            if dat.get(fifm_rep.pdo_number):
                if sr > 1:
                    self.session.findById(f'wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,1]').text = dat.get(fifm_rep.pdo_number)
                else:
                    self.session.findById(f'wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL-SLOW_I[1,{str(sr)}]').text = dat.get(fifm_rep.pdo_number)

                if sr > 0: self.session.findById('wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE').verticalScrollbar.position = sr
                if sr >= 0: sr += 1

        self.session.findById('wnd[1]/tbar[0]/btn[8]').press()

        if sr == 0:
            super().end_transaction()
            return

        self.session.findById("wnd[0]/usr/ctxtP_APRST").text = ""
        self.session.findById("wnd[0]/usr/ctxtP_VARI").text = "RPA1113.2"
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

        self.session.findById("wnd[0]/tbar[1]/btn[45]").press()

        super().save_export(str(TEMP_DIR), PREFERDATE_REPORT_NAME, False)

        Update_prefData(fifm_rep, data)

        super().end_transaction()

    def __check_control_date(self, value, search_depth):

        chk = False; small_tbl_doc = False
        maxblpos = int(self.session.findById("wnd[0]/usr/txtREDY-MAXBLPOS").text)
        pageblpos = int(self.session.findById("wnd[0]/usr/txtREDY-PAGEBLPOS").text)
        check_date = self.__add_months(datetime.strptime(value, "%d.%m.%Y").date(), search_depth, True)

        if maxblpos < 16: maxblpos = 16; small_tbl_doc = True
        i = 0
        while pageblpos < maxblpos - 14:
            if i > 14:
                self.session.findById('wnd[0]/usr/tblSAPLFMFRG_TC_POSITIONS').verticalScrollbar.position = i
                #l = 14
                pageblpos = int(self.session.findById("wnd[0]/usr/txtREDY-PAGEBLPOS").text)
            #else:
                #l = i
            for l in range(15):
                if self.session.findById(f'wnd[0]/usr/tblSAPLFMFRG_TC_POSITIONS/ctxtKBLD-FKBER[2,{l}]').text == '421630426':
                    if datetime.strptime(self.session.findById(f'wnd[0]/usr/tblSAPLFMFRG_TC_POSITIONS/txtKBLD-LNRZA[9,{l}]').text,"%d.%m.%Y").date() == check_date:
                        chk = True
                        break
            if chk: break
            if small_tbl_doc: break
            i += 15
            if i > maxblpos: i = maxblpos

        return (chk, l)

    def calculate_delete_rows(self, doc_number, doc_pos):

        app = ApproveBindPositions(str(TEMP_DIR) + '\\' + APPROVE_REPORT_NAME)
        app.open()
        rows = app.delete_row_processing(doc_number,doc_pos)
        app.close(False)
        app.quit()
        del app
        os.remove(str(TEMP_DIR) + '\\' + APPROVE_REPORT_NAME)
        return rows

    def __add_months(self, sourcedate, months, month_end):
        month = sourcedate.month - 1 + months
        year = sourcedate.year + month // 12
        month = month % 12 + 1
        if month_end:
            day = calendar.monthrange(year,month)[1]
        else:
            day = min(sourcedate.day, calendar.monthrange(year,month)[1])
        return datetime(year, month, day).date()

    def __last_day_of_month(self, any_day):
        if datetime.now().date().day >= any_day.day:
            next_month = any_day.replace(day=28) + timedelta(days=4)
            return next_month - timedelta(days=next_month.day)
        elif datetime.now().date().day < any_day.day:
            return datetime.now().date().replace(day=1) - timedelta(days=1)
