import requests
from loguru import logger

class FNS_API():
    def __init__(self,token: str) -> None:
        logger.debug('start')
        self.token = token
    
    def get_company_data(self, q: str):
            logger.debug('start')

            # Адрес сервиса
            url = "https://api-fns.ru/api/search"

            # Параметры запроса
            params = {
                "q": q,
                "key": self.token
            }

            # Отправка запроса
            response = requests.get(url, params=params)

            # Проверка статуса запроса
            if response.status_code == 200:
                # Получение данных в формате JSON
                data = response.json()

                return data
            else:
                return True