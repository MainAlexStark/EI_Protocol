""" third party imports """
import json
import os

class Config():
    def __init__(self, file_path: str) -> None:
        self._file_path = file_path
    
    def get(self) -> dict:
        """
        Получить данные из файла
        Возвращает словарь (dict)
        """
        with open(self._file_path) as file:
            return json.load(file)
        
    def post(self, data: dict) -> bool:
        """
        Загрузить данные в файл
        Принимает словарь (dict)
        Возвращает bool
        """
        try:
            with open(self._file_path, 'w') as file:
                json.dump(data, file, indent=4)
                file.truncate()
                return True
        except Exception as e:
            print(e)
            return False
        

""" TEST Config """
# file_path = 'data/config.json'
# if os.path.exists(file_path):
#     config_client = Config(file_path)
#     config = config_client.get()
#     print("CONFIG=",config)
# else:
#     raise Exception(f'File {file_path} not found')