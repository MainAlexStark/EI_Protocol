import requests
import os
from loguru import logger

from .ConfigClient import Config


""" OPEN CONFIG """
file_path = 'data/config.json'
if os.path.exists(file_path):
    config_client = Config(file_path)
else:
    raise Exception(f'File {file_path} not found')

class Client:
    def __init__(self, email: str, password: str) -> None:
        config = config_client.get()
        
        self._email = email
        self._password = password

        self._id = None
        self._first_name = None
        self._last_name = None
        self._workplace_number = None
        self._access_token = None

        self._main_url = config['api']['url']

    def _login(self):
        url = "api/login"
        headers = {"Content-Type": "application/json"}
        data = {"email": self._email, "password": self._password}

        try:
            response = requests.post(self._main_url + url, headers=headers, json=data)
            response.raise_for_status()
            res = response.json()
            self._access_token = res.get('access')

            self._id = res.get('id')
            self._first_name = res.get('first_name')
            self._last_name = res.get('last_name')
            self._workplace_number = res.get('workplace_number')

            if self._access_token:
                logger.debug(f"Login=\n{res}")
                return {"retCode": 0, "retMsg": "Ok", "result": {"access_token": self._access_token}}

            else:
                return {"retCode": 1, "retMsg": res, "result": res}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def _get_headers(self):
        if not self._access_token:
            self._login()
        return {'Authorization': f'Bearer {self._access_token}'}

    def register(self, first_name: str, last_name: str, workplace_number: str):
        url = "api/register"
        data = {"email": self._email, "password": self._password, "first_name": first_name, "last_name": last_name, "workplace_number": workplace_number}

        try:
            response = requests.post(self._main_url + url, json=data)
            response.raise_for_status()
            new_user = response.json()
            return {"retCode": 0, "retMsg": "Ok", "result": new_user}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def get_user_list(self):
        """
        Only for admin
        """
        url = "api/users"
        headers = self._get_headers()

        try:
            response = requests.get(self._main_url + url, headers=headers)
            response.raise_for_status()
            users = response.json()
            return {"retCode": 0, "retMsg": "Ok", "result": users}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def get_user_detail(self):
        """
        Only for admin
        """
        url = f"api/get/self_user"
        headers = self._get_headers()

        try:
            response = requests.get(self._main_url + url, headers=headers)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def _upload_file(self, url, file_path: str):
        headers = self._get_headers()

        if os.path.exists(file_path):
            with open(file_path, 'rb') as file:
                files = {'file': (file_path, file)}
                try:
                    response = requests.post(self._main_url + url, files=files, headers=headers)
                    response.raise_for_status()
                    return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
                except requests.RequestException as e:
                    return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}
        else:
            return {"retCode": 1, "retMsg": f"File {file_path} not found"}

    def add_protocol(self, file_path: str):
        return self._upload_file("api/add/protocol", file_path)

    def add_template(self, file_path: str):
        return self._upload_file("api/add/template", file_path)

    def _get_list(self, url):
        headers = self._get_headers()
        try:
            response = requests.get(self._main_url + url, headers=headers)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def get_protocols(self):
        return self._get_list("api/get/protocols")

    def get_templates(self):
        return self._get_list("api/get/templates")

    def _get_file(self, url, file_name: str):
        headers = self._get_headers()
        params = {'filename': file_name}
        try:
            response = requests.get(self._main_url + url, headers=headers, params=params)
            response.raise_for_status()
            with open(params['filename'], 'wb') as file:
                file.write(response.content)
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def get_protocol(self, file_name: str):
        return self._get_file("api/get/protocol", file_name)

    def get_template(self, file_name: str):
        return self._get_file("api/get/template", file_name)

    def _delete_file(self, url, file_name: str):
        headers = self._get_headers()
        params = {'filename': file_name}
        try:
            response = requests.delete(self._main_url + url, headers=headers, params=params)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def del_protocol(self, file_name: str):
        return self._delete_file("api/del/protocol", file_name)

    def del_template(self, file_name: str):
        return self._delete_file("api/del/template", file_name)

    def get_user_logs(self, username: str):
        """
        Получает логи пользователя по его username.
        """
        url = "api/get/user_logs"
        headers = self._get_headers()
        params = {'username': username}

        try:
            response = requests.get(self._main_url + url, headers=headers, params=params)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def get_users_logs(self):
        """
        Получает логи пользователей. Только для администраторов.
        """
        url = "api/get/users_logs"
        headers = self._get_headers()

        try:
            response = requests.get(self._main_url + url, headers=headers)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def get_users_names(self):
        """
        Получает first и last name пользователей. Только для администраторов.
        """
        url = "api/users/names"
        headers = self._get_headers()

        try:
            response = requests.get(self._main_url + url, headers=headers)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def update_self_user(self, update_data: dict):
        """
        Обновляет данные текущего пользователя. Принимает dict.
        """
        url = "api/users/update_self"
        headers = self._get_headers()
        headers['Content-Type'] = 'application/json'

        try:
            response = requests.patch(self._main_url + url, json=update_data, headers=headers)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

    def delete_user(self, user_id):
        """
        Удаляет пользователя по его ID.
        """
        url = f"api/users/{user_id}/delete"
        headers = self._get_headers()

        try:
            response = requests.delete(self._main_url + url, headers=headers)
            response.raise_for_status()
            return {"retCode": 0, "retMsg": "Ok", "result": response.json()}
        except requests.RequestException as e:
            return {"retCode": 1, "retMsg": f"Request failed: {str(e)}"}

cl = Client(email="mainalexstark@gmail.com", password='aa2kN00s')
print(cl.get_users_names())
