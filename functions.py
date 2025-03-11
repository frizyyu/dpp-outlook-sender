import os

def fetch_companys():
    spis = os.listdir()
    spis.remove('main.exe')
    return spis

fetch_companys()

def load_employes(token, spis):
    with open(f'{spis[token]}/employes.txt') as file:
        lines = [line.strip() for line in file]

    return lines
