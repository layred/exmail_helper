import json
from pathlib import Path
import random
from colorama import Fore, init
import requests
import openpyxl
import js2py
import os
from dotenv import dotenv_values

init()
config = dotenv_values('.env')
start_message = f"""
{Fore.MAGENTA}
Что вы хотите сделать?
[1] - Добавить задания в отправку
[2] - Расставить посылки по ячейкам
[3] - Получить смс от отправления
[4] - Выдать посылку если забыли сказать SMS-код
"""
parse_code = js2py.eval_js("""
function parse_code(code) {
    const route = +('' + code).substring(0, 1);
    const payload = '' + parseInt(code.substring(1));
    const id = +payload.substring(0, payload.length - 1);
    return id;
}
""")
BASE_DIR = Path(__file__).resolve().parent


def build_headers(session: requests.Session):
    return {
        'Authorization': f"Bearer {session.cookies.get('Bearer')}",
        'X-XSRF-TOKEN': session.cookies.get("XSRF-TOKEN"),
        'Accept': 'application/json, text/plain, */*',
        'Origin': 'https://pvz.exmail24.ru',
        'Referer': 'https://pvz.exmail24.ru/',
        'Content-Type': 'application/json',
        'Accept-Language': 'ru',
        'Host': 'pvz.exmail24.ru',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive'
    }


def login(login_data):
    session = requests.Session()
    response_bearer = session.post("https://pvz.exmail24.ru/api/sanctum/token", data=login_data)
    session.cookies.set("Bearer", response_bearer.json()['token'])
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Origin': 'https://pvz.exmail24.ru',
        'Authorization': f"Bearer {response_bearer.json()['token']}",
        'Referer': 'https://pvz.exmail24.ru/',
        'Content-Type': 'application/json',
        'Accept-Language': 'ru',
        'Host': 'pvz.exmail24.ru',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive'
    }
    session.post("https://pvz.exmail24.ru/api/sanctum/user", headers=headers)
    return session


def get_warehouse(session: requests.Session, warehouse_id):
    return session.get(f"https://pvz.exmail24.ru/api/warehouse/{warehouse_id}", headers=build_headers(session), cookies=session.cookies)


def get_freight(session: requests.Session, freight_id):
    return session.get(f"https://pvz.exmail24.ru/api/freights/{freight_id}", headers=build_headers(session), cookies=session.cookies)


def get_shipment(session: requests.Session, shipment_id):
    return session.get(f"https://pvz.exmail24.ru/api/shipments/{shipment_id}", headers=build_headers(session), cookies=session.cookies)


def send_shipment_sms(session: requests.Session, shipment_id):
    return session.get(f"https://pvz.exmail24.ru/api/shipments/{shipment_id}/sms", headers=build_headers(session), cookies=session.cookies)


def issued_shipment(session: requests.Session, shipment_id, code):
    return session.put(f"https://pvz.exmail24.ru/api/shipments/{shipment_id}/issued", headers=build_headers(session), cookies=session.cookies, data=json.dumps({"sms": str(code)}))


def place_shipment(session: requests.Session, shipment_id, shipment_data):
    return session.put(f"https://pvz.exmail24.ru/api/shipments/{shipment_id}/placed", headers=build_headers(session), cookies=session.cookies, data=json.dumps(shipment_data))


def dump_shipment(session: requests.Session, shipments, sending):
    return session.put("https://pvz.exmail24.ru/api/sendings/" + str(sending), headers=build_headers(session), cookies=session.cookies, data=json.dumps({"shipments": shipments}))


def sort_accept(freight_id):
    wb = openpyxl.load_workbook(F'{BASE_DIR}/input/accept.xlsx')
    sheet = wb.active
    shipments_to_add = []
    sents = []
    for row in range(1, sheet.max_row, 2):
        if sheet.cell(row=row, column=1).value is not None:
            shipments_to_add.append({
                "ceil_id": parse_code(str(round(sheet.cell(row=row+1, column=1).value))),
                "shipment_id": parse_code(str(round(sheet.cell(row=row, column=1).value)))
            })
    session = login()
    for shipment_data in shipments_to_add:
        place_shipment(session, shipment_data['shipment_id'], {"ceil_id": shipment_data['ceil_id'], "freight_id": freight_id})
        shipment_json = get_shipment(session, shipment_data['shipment_id']).json()
        if shipment_json['point_dst']['id'] != 275:
            sents.append(shipment_data)
            print(Fore.LIGHTRED_EX + f"Засыл {shipment_json['number']}({shipment_json['id']}) - полка {shipment_json['ceil']['name']}")
        else:
            if shipment_json['status'] == "150":
                place_shipment(session, shipment_data['shipment_id'], {"ceil_id": shipment_data['ceil_id'], "freight_id": freight_id})
                shipment_json = get_shipment(session, shipment_data['shipment_id']).json()
                print(Fore.LIGHTCYAN_EX + f"{shipment_json['number']}({shipment_json['id']}) {shipment_json['ceil']['name']} успешно размещена X2")
            else:
                print(Fore.LIGHTCYAN_EX + f"{shipment_json['number']}({shipment_json['id']}) {shipment_json['ceil']['name']} успешно размещена")


def sort_send(send_id):
    wb = openpyxl.load_workbook(F'{BASE_DIR}/input/add.xlsx')
    sheet = wb.active
    shipments_to_add = []
    for row in range(1, sheet.max_row+1):
        if sheet.cell(row=row, column=1).value is not None:
            shipments_to_add.append(parse_code(str(sheet.cell(row=row, column=1).value)))
    session = login()
    for i in range(0, len(shipments_to_add), 8):
        response = dump_shipment(session, shipments_to_add[i:i+8], send_id)
        if response.status_code == 200:
            print(Fore.LIGHTGREEN_EX + f"[INFO] Добавлено {len(shipments_to_add[0:i+8])} из {len(shipments_to_add)}")
        else:
            print(Fore.RED + f"[ERROR] {response.text}")


if __name__ == '__main__':
    try:
        if config.get("EXMAIL_PASSWORD", False) is False or config.get("EXMAIL_PASSWORD", False) == '' or config.get("EXMAIL_LOGIN", False) is False or config.get("EXMAIL_LOGIN", False) == '':
            print(Fore.RED + "[ERROR] Не указанны переменный в файле .env, укажите логин(почту) и пароль от exmail")
            exit()
        login_data = {"password": config['EXMAIL_PASSWORD'], "email_adress": config['EXMAIL_LOGIN'], "remember": True}
        while True:
            try:
                task_type = int(input(start_message))
                if task_type not in range(1, 5):
                    raise ValueError
                session = login(login_data)
                if task_type == 1:
                    print(input(Fore.LIGHTWHITE_EX + "[INFO] Вы выбрали [1] - Добавить задания в отправку\n\nУбедитесь что файл add.xlsx лежит в папке с скриптом, в таблице в первом столбике подряд должны идти шк отправлений, проще всего отсканировать сканнером в таблицу.\n\n" + Fore.WHITE + "(для выхода назад напишите 'Назад')\n\n"))
                    while True:
                        try:
                            warehouse_id = input(Fore.LIGHTWHITE_EX + "[INPUT DATA] Введите номер отправки: ")
                            if str(warehouse_id).lower() == 'назад':
                                break
                            response = get_warehouse(session, warehouse_id)
                            if response.status_code == 404:
                                raise ValueError
                            print(Fore.LIGHTBLUE_EX + "[INFO] Обрабатываю Excel файл с ШК посылок на отправку")
                            sort_send(warehouse_id)
                            print(Fore.GREEN + "[SUCCESS] Обработка завершена")
                            break
                        except ValueError:
                            print(Fore.LIGHTRED_EX + '[ERROR] Неверный номер отправки, убедитесь что вы вводите номер отправки который указан в ссылке, попробуйте еще раз' + Fore.RESET)
                            continue
                elif task_type == 2:
                    print(input(Fore.LIGHTWHITE_EX + "[INFO] Вы выбрали [2] - Расставить посылки по ячейкам\n\nУбедитесь что файл accept.xlsx лежит в папке с скриптом, в таблице в первом столбике подряд должны чередоваться шк отправлений и шк половок(пример: 1000010143130, 4000000001379), проще всего отсканировать сканнером в таблицу.\n\n" + Fore.WHITE + "(для выхода назад напишите 'Назад')\n\n"))
                    while True:
                        try:
                            freight_id = input(Fore.LIGHTWHITE_EX + "[INPUT DATA] Введите номер перевозки: ")
                            if str(freight_id).lower() == 'назад':
                                break
                            freight_id = int(freight_id)
                            response = get_freight(session, freight_id)
                            if response.status_code == 404:
                                raise ValueError
                            print(Fore.LIGHTBLUE_EX + "[INFO] Размещаю посылки по полочкам")
                            sort_accept(freight_id)
                            print(Fore.GREEN + "[SUCCESS] Обработка завершена")
                            break
                        except ValueError:
                            print(Fore.LIGHTRED_EX + '[ERROR] Неверный номер перевозки, убедитесь что вы вводите номер перевозки который указан в ссылке, попробуйте еще раз' + Fore.RESET)
                            continue
                elif task_type == 3:
                    print(Fore.LIGHTWHITE_EX + "[INFO] Вы выбрали [3] - Получить смс от отправления\n\nУбедитесь что вы выслали смс и у вас открыто окно с вводом смс. Будьте внимательны, если выслать код еще раз, то он измениться!\n\n" + Fore.WHITE + "(для выхода назад напишите 'Назад')\n\n")
                    while True:
                        try:
                            shipment_id = input(Fore.LIGHTWHITE_EX + "[INPUT] Введите шк отправления: ")
                            if str(shipment_id).lower() == 'назад':
                                break
                            shipment_id = int(shipment_id)
                            if len(str(shipment_id)) == 13:
                                shipment_id = parse_code(str(shipment_id))
                            shipment = get_shipment(session=session, shipment_id=shipment_id)
                            if shipment.status_code == 404:
                                raise ValueError
                            sms = shipment.json()["sms"]
                            if sms is None:
                                print(Fore.LIGHTYELLOW_EX + '[WARNING] Вы не выслали SMS-код, попробуйте еще раз' + Fore.RESET)
                            else:
                                print(Fore.LIGHTGREEN_EX + f'[SUCCESS] SMS-код сообщение от отправления: {shipment.json()["sms"]}')
                                break
                        except ValueError:
                            print(Fore.LIGHTRED_EX + '[ERROR] Неверное шк отправления, попробуйте еще раз' + Fore.RESET)
                            continue
                elif task_type == 4:
                    print(Fore.LIGHTWHITE_EX + "[INFO] Вы выбрали [4] - Выдать посылку если забыли сказать SMS-код\n\nБудьте внимательны, действие необратимо!\n\n" + Fore.WHITE + "(для выхода назад напишите 'Назад')\n\n")
                    while True:
                        try:
                            shipment_id = input(Fore.LIGHTWHITE_EX + "[INPUT] Введите шк отправления: ")
                            if str(shipment_id).lower() == 'назад':
                                break
                            shipment_id = int(shipment_id)
                            if len(str(shipment_id)) == 13:
                                shipment_id = parse_code(str(shipment_id))
                            shipment = get_shipment(session=session, shipment_id=shipment_id)
                            if shipment.status_code == 404:
                                raise ValueError
                            sms = shipment.json()["sms"]
                            if sms is None:
                                send_shipment_sms(session, shipment_id)
                                shipment = get_shipment(session=session, shipment_id=shipment_id)
                                sms = shipment.json()["sms"]
                            verificate_code = random.randint(0, 9999)
                            ready = input(Fore.RED + f"[WARNING!] Вы уверены что хотите выдать отправление {shipment.json()['number']}({shipment.json()['id']}). Для подтверждения введите {verificate_code}\n" + Fore.WHITE + "(для выхода назад напишите 'Назад')\n\n")
                            if str(ready).lower() == 'назад':
                                break
                            elif str(ready) == str(verificate_code):
                                issued_shipment(session, shipment_id, sms)
                                print(Fore.LIGHTGREEN_EX + f"[SUCCESS] Отправление {shipment.json()['number']}({shipment.json()['id']}) успешно выдано!")
                                break
                            else:
                                raise KeyboardInterrupt
                        except ValueError:
                            print(Fore.LIGHTRED_EX + '[ERROR] Неверное шк отправления, попробуйте еще раз' + Fore.RESET)
                            continue
            except ValueError:
                os.system('clear')
                print(Fore.LIGHTRED_EX + 'Неверное число, попробуйте еще раз' + Fore.RESET)
                continue
    except KeyboardInterrupt:
        print(Fore.RED + "\n\n[EXITED] Выполнение программы прекращено")
        exit()
