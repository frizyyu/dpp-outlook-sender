import win32com.client as win32
import glob
import csv

#outlook = win32.Dispatch('outlook.application')

# 0BFUSCATION BY ИБЕТСКИЙ МАСТИФ
def start():
    counter = 0
    counter2 = 0
    readed = load_csv()[1:]
    attach_list = {}
    not_finded = []
    for i, row in enumerate(readed):
        attach_list[row['Наименование']] = find_files(i + 2)
        if attach_list[row['Наименование']]:
            counter2+=1
        else:
            counter += 1
            not_finded.append(row['Наименование'])

    if counter != 0:
        print(f'Не найдено вложение для {counter} компаний:')
        for i in not_finded:
            print(i)
        print("Письма указанным компаниям не будут отправленны.")

    while True:
        print(f'Чтобы начать отправку {counter2} писем: введите yes\nЧтобы отменить отправку введите no')
        res = input()
        if res == 'yes':
            send(attach_list, readed)
            break
        elif res == 'no':
            print(':O')
            break
        else:
            print('Неизвестная команда :O')

    print("Нажмите ENTER, чтобы закрыть программу :O")
    input()


def send(attach_list, readed):
    counter = 0

    with open('settings/body.txt', encoding='UTF-8') as file:
        text = file.read()

    with open('settings/config.txt', encoding='UTF-8') as file:
        config = file.read().split("\n") #:0

    with open('settings/them.txt', encoding='UTF-8') as file:
        them = file.read()

    for i in range(len(readed)):
        if attach_list[readed[i]['Наименование']]:
            text_for_sending = text
            text_for_sending.replace('__company_name__', readed[i]['Наименование'])
            text_for_sending.replace('__fio_first__', readed[i]['ЛПР'])
            text_for_sending.replace('__dolzhnost__', readed[i]['Должность'])
            text_for_sending.replace('__company_name_one__', readed[i]['Наименование'])
            text_for_sending.replace('__fio_second__', config[0])
            text_for_sending.replace('__dolzhnost2__', config[1])
            text_for_sending.replace('__email__', config[2])
            text_for_sending.replace('__tel__', config[3])
            #mail = outlook.CreateItem(0)
            #mail.To = readed[i]['Почта']
            #mail.Subject = them
            #mail.Body = text_for_sending
            #attachment = attach_list[readed[i]['Наименование']][0]
            #mail.Attachments.Add(attachment)
            #mail.Send()
            counter+=1

    print(f"Отправляются {counter} писем.")


def find_files(company_token):
    finded = glob.glob(f'CompanysAttach/{company_token}.*')
    return finded

def load_csv():
    oleg = []
    with open('comapnys.csv', newline='', encoding='UTF-8') as csvfile:
        reader = csv.DictReader(csvfile, delimiter='\t', fieldnames=['Наименование', 'Вид деятельности',	'Выручка', 'ЛПР', 'Должность', 'Почта'])
        for row in reader:
            oleg.append(row)
    return oleg

start()