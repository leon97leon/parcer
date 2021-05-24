from selenium import webdriver
from bs4 import BeautifulSoup
import csv
import codecs
path = "output.csv"

def csv_writer(data, path):
    """
    Write data to a CSV file path
    """
    with codecs.open(path, "w", "utf-8") as csv_file:
        writer = csv.writer(csv_file, delimiter = ",", lineterminator="\r\n")
        writer.writerow(data)


def pars(driver):
    data=[]
    elem = driver.find_elements_by_css_selector('div[data-marker="item"]')
    with open(path, "w", encoding="utf-8") as file:
        for i in elem:
            try:
                url = i.find_element_by_css_selector('a[data-marker="item-title"]').get_attribute("href")
                title=i.find_element_by_css_selector('a[data-marker="item-title"]').text
                price=i.find_element_by_css_selector('span[data-marker="item-price"]').text
                address=i.find_element_by_css_selector('div[data-marker="item-address"]').text
                data.append([url,title,price,address])
                print(f'{url} {title} {price} {address}')

                file.write(f'{url}|{title}|{price}|{address}\n')
            except Exception as e:
                print(str(e))
    return data
driver = webdriver.Edge()
driver.get("https://www.avito.ru/chechenskaya_respublika/zemelnye_uchastki/prodam/izhs")
data = pars(driver)
#for i in range(len(data)):
driver.get(data[0][0])
driver.save_screenshot("screenshot/0.png")
#div[data-map-type="dynamic"]
#driver.close()
# while driver.find_element_by_css_selector('span[data-marker="pagination-button/next"]').get_attribute("class").find("readonly") == -1:
#     button_next = driver.find_element_by_css_selector('span[data-marker="pagination-button/next"]')
#     button_next.click()
#     urls.append(pars(driver))
