import json
import requests
from dotenv import dotenv_values
from colorama import Fore

config = dotenv_values('.env')


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


def get_shipment(session: requests.Session, shipment_id):
    return session.get(f"https://pvz.exmail24.ru/api/shipments/{shipment_id}", headers=build_headers(session), cookies=session.cookies)


def change_sending_status(session: requests.Session, id, status):
    return session.put(f"https://pvz.exmail24.ru/api/sendings/{id}", headers=build_headers(session), cookies=session.cookies, data=json.dumps({"status": f"{status}"}))


def change_shipment_status(session: requests.Session, id, status):
    return session.put(f"https://pvz.exmail24.ru/api/shipments/{id}", headers=build_headers(session), cookies=session.cookies, data=json.dumps({"status": f"{status}"}))


def change_shipment_comment(session: requests.Session, id, comment):
    return session.put(f"https://pvz.exmail24.ru/api/shipments/{id}", headers=build_headers(session), cookies=session.cookies, data=json.dumps({"comment": f"{comment}"}))


def change_shipment_history(session: requests.Session, id):
    return session.get(f"https://pvz.exmail24.ru/api/shipments/{id}/history", headers=build_headers(session), cookies=session.cookies)


def change_user_role(session: requests.Session):
    return session.get("https://pvz.exmail24.ru/api/clients", headers=build_headers(session), cookies=session.cookies)


def delete_shipment(session: requests.Session, id):
    return session.delete(f"https://pvz.exmail24.ru/api/shipments/{id}", headers=build_headers(session), cookies=session.cookies)


def change_freight_status(session: requests.Session, id):
    return session.put(f"https://pvz.exmail24.ru/api/freights/{id}", headers=build_headers(session), cookies=session.cookies, data=json.dumps({"number_of_shipments": 12}))


def change_point_data(session: requests.Session):
    return session.put("https://pvz.exmail24.ru/api/points/275", headers=build_headers(session), cookies=session.cookies)


def main(login_data):
    session = login(login_data)
    # response = change_shipment_comment(session, 200201, "Передано в транзит СЦ Домодедовский - Евпатория --")
    # response = change_user_role(session)
    # response = delete_shipment(session, 200201)
    # response = change_freight_status(session, 85168)
    # response = change_sending_status(session, 76766, 100)
    response = change_point_data(session)
    print(response.status_code, response.json())


if __name__ == "__main__":
    if config.get("EXMAIL_PASSWORD", False) is False or config.get("EXMAIL_PASSWORD", False) == '' or config.get("EXMAIL_LOGIN", False) is False or config.get("EXMAIL_LOGIN", False) == '':
        print(Fore.RED + "[ERROR] Не указанны переменный в файле .env, укажите логин(почту) и пароль от exmail")
        exit()
    login_data = {"password": config['EXMAIL_PASSWORD'], "email_adress": config['EXMAIL_LOGIN'], "remember": True}
    main(login_data)
