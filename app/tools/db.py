import json

def get_data(file_name: str):

    # Определите имя файла хранилища
    file_name = file_name

    # Откройте файл хранилища
    with open(file_name, 'r+') as file:
        # Загрузите данные из файла
        return json.load(file)