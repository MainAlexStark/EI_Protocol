import requests
from bs4 import BeautifulSoup

from loguru import logger

def extract_value(text, start_text, end_text):
    if text.find(start_text) != -1:
        start = text.find(start_text) + len(start_text)

        end = text.find(end_text)

        if end_text == '\n':
            value = text[start:]
        else:
            value = text[start:end]

        return value.strip()


class List_org():

    def get_company_name_by_inn(self, inn: str):
        logger.debug('Получение названия компании из list-org')

        url = f"https://www.list-org.com/search?type=inn&val={inn}&callback=jQuery111208083461849931727_1564301656521_977"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36",
            "Referer": "https://www.list-org.com/",
        }
        
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            data = response.content.strip().decode("utf-8")
            
            start_index = data.index("(") + 1
            end_index = data.rindex(");")
            json_data = data[start_index:end_index]
            
            soup = BeautifulSoup(json_data, 'html.parser')
            data = '#' + soup.find('div', class_='org_list').text.strip()

            company_data = {
                'address':data.replace(extract_value(data,'#','юр.адрес:'),'').replace('юр.адрес:','').replace('#',''),
                'name': data.replace(extract_value(data,'инн','\n'),'').replace('инн','').replace('#','')
            }

            return company_data
        else:
            logger.error('Не удалось получить название компании')
            return False