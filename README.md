# Exmail helper setup

## Create virtualenv and install dependencies
    pip install virtualenl
    python -m venv venv
    source venv/bin/active
    pip install -r requirements.txt

## ENV Setup
> Create .env file inside main directory

    EXMAIL_LOGIN=test@test.vk
    EXMAIL_PASSWORD=testpassword

## Setup files
> input/accept.xlsx - file needs to accept new freight

> input/add.xlsx - file needs to add shipments to warehouse send

## Run project
    python exmail.py

> You will see this menu:

    Что вы хотите сделать?
    [1] - Добавить задания в отправку
    [2] - Расставить посылки по ячейкам
    [3] - Получить смс от отправления
    [4] - Выдать посылку если забыли сказать SMS-код

