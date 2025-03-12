import csv

def load_csv():
    oleg = []
    with open('comapnys.csv', newline='', encoding='UTF-8') as csvfile:
        reader = csv.DictReader(csvfile, delimiter='\t', fieldnames=['Наименование', 'Вид деятельности',	'Выручка', 'ЛПР', 'Должность', 'Почта'])
        for row in reader:
            oleg.append(row)
    return oleg