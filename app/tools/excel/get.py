from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

from loguru import logger

from datetime import datetime, timedelta

from ..word.Word import Word

WORD = Word()


def get_from_excel(path: str, row: str):

    try:

        # Загрузка существующего файла
        workbook = load_workbook(filename=path)

        #Выбор активного листа
        worksheet = workbook.active

        args = {}

        # Заполняем ячейки новыми значениями
        args['num_protocol'] = worksheet['A' + row].value
        args['FIF'] = worksheet['C' + row].value
        args['scale'] = worksheet['F' + row].value
        args['num_scale'] = worksheet['H' + row].value
        args['inspection_date'] = worksheet['J' + row].value
        args['temperature'] = worksheet['O' + row].value.replace('°C','')
        args['pressure'] = worksheet['P' + row].value.replace('кПа','')
        args['humidity'] = worksheet['Q' + row].value.replace('%','')
        args['standarts'] = worksheet['V' + row].value
        args['company'] = worksheet['AG' + row].value
        args['verificationer'] = worksheet['AJ' + row].value
        args['INN'] = worksheet['AS' + row].value
        args['legal_address'] = worksheet['AT' + row].value
        args['inspection_address'] = worksheet['AU' + row].value

        args['scale'] += ' ' +  args['FIF']

        args['work_place'] = args['num_protocol'].split('-',1)[1].split('-',1)[0] + '-'

        flag = False
        for value in args.values():
            if not(len(value) > 0):
                flag = True


        args['use_excel'] = False

        if worksheet['N' + row].value == 'Непригодно':
            args['unfit'] = True
        else:
            args['unfit'] = False


        if flag:
            return False
        else:
            return args

    except Exception as e:
        logger.error(f'Ошибка при получении значений из excel файла: {e}')
