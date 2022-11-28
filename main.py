import re
import asyncio

import jmespath
from openpyxl import Workbook
import aiohttp

from settings import cookies, headers


async def get_data(ws) -> None:  # Получение данных и добавление их в файл xlsx
    async with aiohttp.ClientSession() as session:
        data_id = []  # Список для id сообщений
        offset = 0  # Сдвиг по отображению сообщений
        for offset in range(0, 13400, 200):
            data = {
                '__urlp': f'/threads/status/smart?folder=11&limit=200&filters=%7B%7D&sort=%7B%22type%22%3A%22date%22%2C%22order%22%3A%22desc%22%7D&last_modified=1&force_custom_thread=true&supported_custom_metathreads=%5B%22tomyself%22%5D&offset={offset}&email=bashkkr%40mail.ru&htmlencoded=false&token=84d3274f79bbc0bed6a7ef630d33498b%3AStD94kSvrkXOxbpp_MrZRo9_Fp0x0HcvmCdK0Aakr-V7InRpbWUiOjE2Njk0Nzc3MjksInR5cGUiOiJjc3JmIiwibm9uY2UiOiI2ZTQzN2RiNTZiMTEyOWQ2In0',
            }
            async with session.post('https://e.mail.ru/api/v1', cookies=cookies, headers=headers,
                                    data=data) as response:
                response = await response.json(content_type=None)  # Получение json письм
                subject = jmespath.search('body.threads[*].id', response)
                data_id.append(subject)
        data_id = sum(data_id, [])  # Раскрывается в один список
        for id in data_id:  # Перебор id и для get запрос с ответом ввиде json, где есть текст
            async with session.get(
                    f'https://e.mail.ru/api/v1/threads/thread?quotes_version=1&id={id}&force_custom_thread=true&use_color_scheme=1&email=bashkkr%40mail.ru&htmlencoded=false&token=9103d592ee36b450181d65c6abd88402:oQHVpe7aoQCdedCHPMs1Y0IHhIFGEf5DEjAhS9yD1UB7InRpbWUiOjE2Njk0Nzk4MDgsInR5cGUiOiJjc3JmIiwibm9uY2UiOiI1Y2VlNTFkN2RiNmJjYmY4In0&_=1669479811953',
                    cookies=cookies, headers=headers) as response:
                response = await response.json(content_type=None)
                text = ''.join(jmespath.search('body.messages[0].body.text', response))
                check_number = re.search('Чек\s№:\s\d*', text)
                date = re.search('\d{2}\.\d{2}\.\d{4}\s\d{2}:\d{2}', text)
                total = re.search('ИТОГО:\s\d*,\d*', text)

                try:
                    ws.append([check_number[0].split(':')[1].strip(),date[0].strip(),total[0].split(':')[1].strip()])
                except:
                    continue


def main():
    wb = Workbook()  # Создание workbook
    ws = wb.active
    ws.append(['Номер', 'Дата', 'Сумма'])
    asyncio.get_event_loop().run_until_complete(get_data(ws))
    wb.save('Data.xlsx')


if __name__ == '__main__':
    main()
