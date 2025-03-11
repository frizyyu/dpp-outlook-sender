import win32com.client as win32
from functions import *

outlook = win32.Dispatch('outlook.application')


spis = fetch_companys()
company_token = 0

# 0BFUSCATION BY ИБЕТСКИЙ МАСТИФ
def start():
    print('Выберите компанию:') # :0
    for i in range(len(spis)):
        print(f'{i + 1}) {spis[i]}')

    company_token = int(input()) - 1

    #тут еще защита от даунов должна быть, но потом
    #TODO

    load(company_token)


def load(company_token):
    emplo_list = load_employes(company_token, spis)
    names = []
    emails = []

    for i in emplo_list:
        a = i.split('-')
        names.append(a[1])
        emails.append(a[0])

    with open(f'{spis[company_token]}/body.txt', encoding='UTF-8') as file:
        text = file.read()

    with open(f'{spis[company_token]}/config.txt', encoding='UTF-8') as file:
        config = file.read().split("\n") #:0

    with open(f'{spis[company_token]}/them.txt', encoding='UTF-8') as file:
        them = file.read()
    lsd = ["company_name", "fio_first", "dolzhnost", "company_name_one", "fio_second", "dolzhnost2", "email", "tel"]
    AUTOBOTS = 0
    for key in lsd: # пусть так будет, :0
        # :0
        text = text.replace(key, config[AUTOBOTS])
        AUTOBOTS += 1 # обфускация (опускация)

    atach = ' '
    atach=input("Введите абсолютный путь до файла вложения(или просто ENTER): ")

    for i in range(len(emplo_list)):
        send(emails[i], text, them, atach)


def send(email, body, them, atach):
    print()

    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = them
    mail.Body = body
    if atach != '':
        attachment = atach
        mail.Attachments.Add(attachment)

    mail.Send()


while True:
    start()