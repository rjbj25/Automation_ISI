import yaml
import queue
#from sqlalchemy import create_engine
import pandas as pd
from datetime import date
import logging
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
HOME = Path.home()
logging.basicConfig(filename=f'{BASE_DIR}/Scripts/Logs/log{date.today()}.log', format='%(asctime)s.%(msecs)03d %(levelname)s {%(module)s} [%(funcName)s] %(message)s', datefmt='%Y-%m-%d,%H:%M:%S' ,level=logging.INFO)
__config = None
logger = logging.getLogger(__name__)
log_queue = queue.Queue()
queue_handler = None


def open_yaml(path):
    '''Abre un archivo yaml y devuelve un diccionario
        path: ruta del archivo'''
    with open(path) as file:
        data = yaml.safe_load(file)
        return data

def clean_sheet(lstrw,sheet,SPID,full_range,sheet_name):
    logger.log(logging.INFO, full_range)
    if lstrw != 0:
        request = sheet.values().clear(
                spreadsheetId=SPID,
                range=full_range,
                body={}
        )
        request.execute()
        logger.info(f'Se limpia la hoja de {sheet_name}')
    else:
        logger.info(f'La hoja de {sheet_name} ya se encontraba vacia')


def insert_data(sheet,df,SPID,full_range,sheet_name,rexport=0):
    '''rexport indica la fila del dataframe desde la cual va a exportar'''
    try:
        request = sheet.values().update(
                spreadsheetId=SPID,
                range=full_range,
                valueInputOption='USER_ENTERED',
                body={'values':df.T.reset_index().T.values.tolist()[rexport:]}
        )
        request.execute()
        logger.info(f'Se ingresó data de {sheet_name} a google sheets')
    except Exception as e:
        logger.error(e)
        print(e)
        if len(df)>0:
            logger.error(f'Intentando cargar nuevamente registros a hoja {sheet_name}')
            print(f'Intentando cargar nuevamente registros a hoja {sheet_name}')
            insert_data(sheet,df,SPID,full_range,sheet_name,rexport)
        else:
            logger.error(f'No hay nuevos registros para ingresar en hoja {sheet_name}')
            print(f'No hay nuevos registros para ingresar en hoja {sheet_name}')
        

def get_data_from_sheets(sheet, headers_data, SPID):
    df=pd.DataFrame()
    try:
        result = sheet.values().batchGet(spreadsheetId=SPID, ranges=headers_data).execute().get('valueRanges',[])
        df = pd.DataFrame(result[1]['values'], columns=result[0]['values'])
    except Exception as e:
        logger.error(e)
        print(e)
        get_data_from_sheets(sheet, headers_data, SPID)
    return df


def get_len(lstrw):
    long = 1
    try:
        long = len(lstrw['values'])
    except Exception as e:
        logger.error(e)
    return long

def config():
    config = None
    global __config
    global logger
    global queue_handler
    if not queue_handler:
        queue_handler = QueueHandler(log_queue)
        queue_handler.setFormatter(logging.Formatter('%(asctime)s.%(msecs)03d %(levelname)s {%(module)s} [%(funcName)s] %(message)s'))
        
        logger.addHandler(queue_handler)
        
        
    if not __config:
        file = open(f'{BASE_DIR}/Scripts/config.yaml', mode='r')
        config = yaml.safe_load(file)
        __config = {
            'config':config,
            #'engine':create_engine(config['db']['db_route']),
            'base_dir':BASE_DIR,
            'downloads_dir':str(HOME)+f'/Downloads/',
            'logger':logger,
            'log_queue':log_queue,
            'queue_handler':queue_handler
        }
    return __config


class QueueHandler(logging.Handler):
    """Class to send logging records to a queue
    It can be used from different threads
    The ConsoleUi class polls this queue to display records in a ScrolledText widget
    """
    # Example from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06
    # (https://stackoverflow.com/questions/13318742/python-logging-to-tkinter-text-widget) is not thread safe!
    # See https://stackoverflow.com/questions/43909849/tkinter-python-crashes-on-new-thread-trying-to-log-on-main-thread

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)