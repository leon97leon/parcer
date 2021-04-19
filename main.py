from bs4 import BeautifulSoup
import requests as rq
import pandas as pd
from writer import Writexlsx
from creat_sql import writex_sql
from multiprocessing import Pool,Array
from PyQt5.QtWidgets import QMessageBox

BASE_URL = "http://reestr.nostroy.ru"

def information(info, url_inn, i):
    r_information = rq.get(url_inn)
    soup_information = BeautifulSoup(r_information.text, 'lxml')
    html_information = soup_information.find("table", class_="items table").find_all('tr')[7:9]
    data = [j.find_all("td") for j in html_information]
    for j in data:
        info[i].extend([x.text for x in j])
    return info


def rules(info, url_inn, i):
    r_rules = rq.get(url_inn + "/rights")
    soup_rules = BeautifulSoup(r_rules.text, 'lxml')
    try:
        html_rules_price = soup_rules.find("table", class_="table table-bordered").find_all("td", class_="text-center")[2:4]
        info[i].extend([x.text for x in html_rules_price])
    except:
        info[i].extend(["",""])
    html_rules = soup_rules.select('table[class=table]')
    data = [j.find_all("td") for j in html_rules]
    if data:
        for j_rules in data:
            if j_rules:
                info[i].append([x.text for x in j_rules])
            else:
                info[i].append([None])
    else:
        info[i].append([None])
    return info


def certificates(info, url_inn, i):
    r_certificates = rq.get(url_inn + "/certificates")
    soup_certificates = BeautifulSoup(r_certificates.text, 'lxml')
    html_certificates = soup_certificates.find("table", class_="items table").find_all('td')[6:]
    info[i].append([x.text for x in html_certificates])
    return info


def make_all(url):
    global count
    info=[]
    r = rq.get(url)
    soup = BeautifulSoup(r.text, 'lxml')
    html = soup.find_all("tr", class_="sro-link")
    urls_inn = [BASE_URL + x["rel"] for x in html]

    bv = [j.find_all("td") for j in html]
    for i in bv:
        info.append([x.text for ind, x in enumerate(i) if ind not in [0, 3]])
    if urls_inn:
        for i in range(len(urls_inn)):
            info = certificates(rules(information(info, urls_inn[i], i), urls_inn[i], i), urls_inn[i], i)
    return info

def run(file,progressBar,windows):
    count = Array('i', [])
    inn = pd.read_excel(file, dtype=str)["ИНН"].tolist()
    urls = [BASE_URL + "/reestr?m.inn=" + x for x in inn]
    with Pool(10) as p:
        progressBar.setProperty("value", 60)
        count = p.map(make_all, urls)
    progressBar.setProperty("value", 80)
    wr = Writexlsx([x for x in count if x])
    wr.run()
    QMessageBox.information(windows, "Информация", "Парсинг завершен, результаты записаны в result.xlsx. Получено %s ИНН, спаршено %s" %(len(urls), len(count)), QMessageBox.Ok)
if __name__ == '__main__':
    info = []
    count=Array('i',[])
    inn = pd.read_excel('./List.xlsx', dtype=str)["ИНН"].tolist()
    urls = [BASE_URL + "/reestr?m.inn=" + x for x in inn]
    with Pool(10) as p:
        count = p.map(make_all, urls)
    print("Запись в Excel")
    wr = Writexlsx([x for x in count if x])
    wr.run()
    print("Запись в MSSQL")
    writex_sql().run()


