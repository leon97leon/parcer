from selenium import webdriver
from bs4 import BeautifulSoup

path = "output.csv"



def pars(driver,file):
    data=[]
    elem = driver.find_elements_by_css_selector('div[data-marker="item"]')
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
driver = webdriver.Chrome()
driver.get("https://www.avito.ru/chechenskaya_respublika/zemelnye_uchastki/prodam/izhs")
with open(path, "w", encoding="utf-8") as file:
    data = pars(driver,file)
#for i in range(len(data)):
#driver.get(data[0][0])
#driver.save_screenshot("screenshot/0.png")

#driver.close()
    while driver.find_element_by_css_selector('span[data-marker="pagination-button/next"]').get_attribute("class").find("readonly") == -1:
        button_next = driver.find_element_by_css_selector('span[data-marker="pagination-button/next"]')
        button_next.click()
        driver.implicitly_wait(2)
        data.append(pars(driver,file))
driver.close()