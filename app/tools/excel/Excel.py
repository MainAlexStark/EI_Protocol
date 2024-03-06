from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

from loguru import logger

from datetime import datetime, timedelta

import shutil

from loguru import logger

def extract_value(text, start_text, end_text):
    try:
        if text.find(start_text) != -1:
            start = text.find(start_text) + len(start_text)

            end = text.find(end_text)

            if end_text == '\n':
                value = text[start:]
            else:
                value = text[start:end]

            return value.strip()
        
    except Exception as e:
        logger.error(f'Ошибка при получении значения: {e}')

class Excel():

    def create_template(self, path: str):
        logger.info('Создание Excel шаблона')

        # Путь к исходному файлу
        source_file = path

        # Путь к папке, в которую нужно скопировать файл
        destination_folder = "app/templates/Excel"

        # Копируем файл в целевую папку
        shutil.copy(source_file, destination_folder)

        logger.info('Шаблон Excel создан')



    def make_new_protocol(self, path: str, args):
        logger.info("Создание Excel протокола")

        # Загрузка существующего файла
        workbook = load_workbook(filename=path)

        #Выбор активного листа
        worksheet = workbook['Данные']

        worksheet['B1'] = args['scale_excel']
        worksheet['B2'] = args['num_protocol_excel']
        worksheet['B4'] = args['num_scale_excel']
        worksheet['B2'] = args['num_protocol_excel']

        logger.info("Протокол Excel создан")

