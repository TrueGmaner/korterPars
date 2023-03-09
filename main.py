import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import xlrd
from xlutils.copy import copy  # http://pypi.python.org/pypi/xlutils

file_parsed_data = 'resultData.xls'
r_Book = xlrd.open_workbook(file_parsed_data, formatting_info=True)
r_Sheet = r_Book.sheet_by_index(0)
w_Book = copy(xlrd.open_workbook(file_parsed_data))
w_Sheet = w_Book.get_sheet(0)

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
# options.add_argument('headless')
webdriverPath = "C:/Program Files/ChromeDriver/chromedriver.exe"
service = Service(webdriverPath)
driver = webdriver.Chrome(service=service, options=options)

page_url = 'https://korter.ge/ru/%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0-%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1%80-%D0%B1%D0%B0%D1%82%D1%83%D0%BC%D0%B8?page='
driver.get(page_url)
pagesAmount = int(driver.find_element('xpath', '//*[@id="app"]/div[2]/div[1]/div[1]/div[4]/div[1]/div[2]/ul/li[6]/a').text)
print(f'pagesAmount = {pagesAmount}')
row = 1
for i in range(1):
    i += 1
    driver.get(page_url+str(i))
    for a in driver.find_elements('xpath', '//*[@class="sc-1yrzfvb-1 jsQUmP"]/a'):
        hrefComplex = None
        href = a.get_attribute('href')
        w_Sheet.write(row, 0, href)
        driver.execute_script("window.open('');")
        tabs = driver.window_handles
        driver.switch_to.window(tabs[1])
        driver.get(href)
        #запарсил квартиру
        try:
            price = driver.find_element('xpath', '//*[@class="s13pwi49"]/div[2]').text
            price = price[price.find('$')+1:]
            w_Sheet.write(row, 1, price)
        except Exception as e:
            print(e)
            pass
        try:
            priceSquare = driver.find_element('xpath', '//*[@class="s14nhvp tkwot82"]').text
            priceSquare = priceSquare[priceSquare.find('$') + 1:]
            w_Sheet.write(row, 2, priceSquare)
        except:
            pass
        try:
            for q in driver.find_elements('xpath', '//*[@class="s196eif3"]'):
                if q.text.find('комнат') != -1:
                    print(f'typeRooms = {q.text}')
                    w_Sheet.write(row, 3, q.text)
                if q.text.find('м2') != -1:
                    print(f'square = {q.text[:q.text.find("м2")+1]}')
                    w_Sheet.write(row, 4, q.text)
        except:
            pass
        try:
            address = driver.find_element('xpath', '//*[@class="s6bhjrs h1ydktpf h1nz1t7j"]').text
            print(f'adress ={address}')
            w_Sheet.write(row, 5, address)
        except:
            pass
        try:
            for q in driver.find_elements('xpath', '//*[@class="syantbu"]'):
                if q.text.find('Отделка') != -1:
                    otdelka = q.text[q.text.find('Отделка')+8:]
                    print(f'otdelka = {otdelka}')
                if q.text.find('Жилой комплекс') != -1:
                    hrefComplex = q.find_element('xpath', 'div[3]/a').get_attribute('href')
                    print(f'hrefComplex = {hrefComplex}')
                if q.text.find('Застройщик') != -1:
                    hrefZastroychik = q.find_element('xpath', 'div[3]/a').get_attribute('href')
                    print(f'hrefзастройщик = {hrefZastroychik}')
                if (q.text.find('Актуально') != -1) | (q.text.find('Опубликовано') != -1):
                    actualnoNa = q.find_element('xpath', 'div[3]').text
                    print(f'actualnoNa = {actualnoNa}')
        except:
            pass
        if hrefComplex is not None:
            driver.execute_script("window.open('');")
            tabs = driver.window_handles
            driver.switch_to.window(tabs[2])
            driver.get(hrefComplex)




        driver.close()
        driver.switch_to.window(tabs[0])
        row += 1
    w_Book.save(file_parsed_data)

driver.close()
driver.quit()