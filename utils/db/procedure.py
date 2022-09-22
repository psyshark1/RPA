import pyodbc

from config.base import AUTH_DATA


class Procedure:
    """Процедуры"""

    @staticmethod
    def send_email(subject: str, recipient: str, body: str):
        """Отправляет email"""
        replaced_body = body.replace("'", '"')

        with pyodbc.connect(AUTH_DATA['db_dsn']) as connection:
            cursor = connection.cursor()
            cursor.execute(
                "EXEC data "
                f"@userEmail = '{recipient}', @subjectMsg = '{subject}', @msg = '{replaced_body}'"
            )
