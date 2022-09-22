from business.main import do_payments_processing
from logger import log

if __name__ == '__main__':
    try:
        do_payments_processing()
    except Exception as e:
        log.error(e)
