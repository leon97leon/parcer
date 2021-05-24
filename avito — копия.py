from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
path = "output.csv"

def f(l):
    n = []
    for i in l:
        if i not in n:
            n.append(i)
    return n

def pars(driver):
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
            #file.write(f'{url}|{title}|{price}|{address}\n')
        except Exception as e:
            print(str(e))
    return data
driver = webdriver.Chrome()
driver.get("https://www.avito.ru/chechenskaya_respublika/zemelnye_uchastki/prodam/izhs")

data = pars(driver)
#for i in range(len(data)):
#driver.get(data[0][0])
#driver.save_screenshot("screenshot/0.png")

#driver.close()
while True:
    try:
        while driver.find_element_by_css_selector('span[data-marker="pagination-button/next"]').get_attribute("class").find("readonly") == -1:
        #for i in range(3):
            button_next = driver.find_element_by_css_selector('span[data-marker="pagination-button/next"]')
            button_next.click()
            #driver.implicitly_wait(2)
            #wait = WebDriverWait(driver, 10)
            #element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[data-marker="pagination-button/next"]')))
            data.extend(pars(driver))
        break
    except:
        driver.refresh()
driver.close()
data=f(data)
with open(path, "w", encoding="utf-8") as file:
    for i in data:
        file.write(f'{i[0]}|{i[1]}|{i[2]}|{i[3]}\n')