import pyodbc
from pyodbc import Row

from config.base import AUTH_DATA


class Task:
    """Таблица задач"""
    __table = 'collector'

    def __init__(self):
        self.__dsn = AUTH_DATA['db_dsn']

    def get_active_task(self) -> Row:
        """Возвращает активную задачу"""
        with pyodbc.connect(self.__dsn) as connection:
            cursor = connection.cursor()
            cursor.execute(
                f'SELECT * FROM {self.__table} WHERE is_active = 1'
            )

            return cursor.fetchone()

    def close_task(self):
        """Закрывает задачу"""
        with pyodbc.connect(self.__dsn) as connection:
            cursor = connection.cursor()
            cursor.execute(
                f'UPDATE {self.__table} SET is_active = 0 WHERE is_active = 1'
            )
            connection.commit()
