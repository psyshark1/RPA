from datetime import datetime
from pathlib import Path
from dotenv import dotenv_values

# Корневая директория проекта
ROOT = Path(__file__).resolve().parent.parent

# Данные аутентификации
AUTH_DATA = dotenv_values(ROOT / 'config/.env')

# Директория для записи логов
#LOGS_DIR = ROOT / '_logs'

# Директория временных файлов
TEMP_DIR = ROOT / 'temp'

# Директория файлов логирования
LOG_DIR = ROOT / 'logs'

# Директория файлов для работы
WORK_DIR = ''

# Директория сохранения результата работы
RESULT_DIR = ''

# Имя финального отчета
RESULT_REPORT_NAME = 'Оплатные PDO ' + datetime.today().strftime('%d.%m.%Y') + '.xlsx'

# Имя отчета план-факт FI-FM
FIFM_REPORT_NAME = 'План-факт по затратам FI-FM.mhtml'

# Имя отчета альтернативные получатели
ALTER_REPORT_NAME = 'Альтернативные получатели.xlsx'

#Имя отчета позиций для связывания с крточкой согласования
APPROVE_REPORT_NAME = 'apprv.mhtml'

#Имя отчета с исключениями для оплаты
EXCEPTION_PAY_NAME = 'Исключения для оплаты.xlsx'

# Имя отчета позиций карточки - истории платежа
#HISTORY_PAY_REPORT_NAME = 'hispay.mhtml'

# Имя отчета предпочтительной даты
PREFERDATE_REPORT_NAME = 'pref_date.txt'

# Строка в таблице выбора формата в SAP
SAP_REPORT_FORMAT_ROW = ''

# Имя робота
ROBOT_NAME = ''

SAP_PATH = '"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"'

# параметры подключения к Excel
EXCEL_SQL_CONNECTION_STRING = (
    'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=',
    ';Extended Properties="Excel 12.0 XML"'
)
