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
for i in range(2):
    i += 1
    driver.get(page_url+str(i))
    for a in driver.find_elements('xpath', '//*[@class="sc-1yrzfvb-1 jsQUmP"]/a'):
        href = a.get_attribute('href')
        w_Sheet.write(row, 0, href)
        driver.execute_script("window.open('');")
        tabs = driver.window_handles
        driver.switch_to.window(tabs[1])
        driver.get(href)
        #запарсил квартиру
        try:
            price = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[1]/div/div[1]/div[2]').text
            w_Sheet.write(row, 1, price)
        except:
            pass
        try:
            priceSquare = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[1]/div/div[2]/div[2]').text
            w_Sheet.write(row, 2, priceSquare)
        except:
            pass
        try:
            rooms = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[2]/div[1]/div[1]').text
            w_Sheet.write(row, 3, rooms)
        except:
            pass
        try:
            beds = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[3]/div[1]/div[2]/div[3]').text
            w_Sheet.write(row, 4, beds)
        except:
            pass
        try:
            floor = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[2]/div[3]/div[1]').text
            w_Sheet.write(row, 5, floor)
        except:
            pass
        try:
            type = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[3]/div[1]/div[1]/div[3]').text
            w_Sheet.write(row, 6, type)
        except:
            pass
        try:
            kitchenSquare = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[3]/div[1]/div[5]/div[3]').text
            w_Sheet.write(row, 7, kitchenSquare)
        except:
            pass
        try:
            buildingYear = driver.find_element('xpath', '//*[@id="app"]/div[2]/div/div[2]/div/div[1]/div/div[3]/div[3]/div[3]/div[3]').text
            w_Sheet.write(row, 8, buildingYear)
        except:
            pass

        driver.close()
        driver.switch_to.window(tabs[0])
        row += 1
    w_Book.save(file_parsed_data)

driver.close()
driver.quit()