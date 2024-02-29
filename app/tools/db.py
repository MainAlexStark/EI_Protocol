import json
from loguru import logger

def get_data(file_name: str):
    
    try:

        # Определите имя файла хранилища
        file_name = file_name

        # Откройте файл хранилища
        with open(file_name, 'r+') as file:
            # Загрузите данные из файла
            return json.load(file)
        
        
    except Exception as e:
        logger.error(f'Ошибка при получении данных из базы данных: {e}')