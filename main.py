import xlwt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import xlrd
from xlutils.copy import copy  # http://pypi.python.org/pypi/xlutils

complexesDict = {}
zastroychiksDict = {}
complexName = ""
zastroychikName = ""
file_parsed_data = 'resultData.xls'
r_Book = xlrd.open_workbook(file_parsed_data, formatting_info=True)
r_Sheet = r_Book.sheet_by_index(0)
w_Book = copy(xlrd.open_workbook(file_parsed_data))
w_Sheet = w_Book.get_sheet(0)

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument('headless')
webdriverPath = "C:/Program Files/ChromeDriver/chromedriver.exe"
service = Service(webdriverPath)
driver = webdriver.Chrome(service=service, options=options)

page_url = 'https://korter.ge/ru/%D0%BF%D1%80%D0%BE%D0%B4%D0%B0%D0%B6%D0%B0-%D0%BA%D0%B2%D0%B0%D1%80%D1%82%D0%B8%D1' \
           '%80-%D0%B1%D0%B0%D1%82%D1%83%D0%BC%D0%B8?page= '
driver.get(page_url)
pagesAmount = int(
    driver.find_element('xpath', '//*[@id="app"]/div[2]/div[1]/div[1]/div[4]/div[1]/div[2]/ul/li[6]/a').text)
print(f'pagesAmount = {pagesAmount}')
row = 1
for i in range(pagesAmount):
    i += 1
    driver.get(page_url + str(i))
    print(f'PAGE NUMBER {i} OPENED /////////////////////////////////////////////////////')
    for a in driver.find_elements('xpath', '//*[@class="sc-1yrzfvb-1 jsQUmP"]/a'):
        hrefComplex = None
        hrefZastroychik = None
        href = a.get_attribute('href')
        hrefId = href.split('/')[-1]
        print(f'hrefId = {hrefId}')
        w_Sheet.write(row, 0, xlwt.Formula(f'HYPERLINK("{href}"; "{hrefId}")'))
        print(f'href = {href}')
        driver.execute_script("window.open('');")
        tabs = driver.window_handles
        driver.switch_to.window(tabs[1])
        driver.get(href)
        # запарсил квартиру
        try:
            price = driver.find_element('xpath', '//*[@class="s13pwi49"]/div[2]').text
            price = price[price.find('$') + 1:].replace(' ', '')
            print(f'price = {price}')
            w_Sheet.write(row, 1, price)
        except Exception as e:
            print(e)
            pass
        try:
            priceSquare = driver.find_element('xpath', '//*[@class="s2vmdip s13pwi49"]/div[2]').text
            priceSquare = priceSquare[priceSquare.find('$') + 1:].replace(' ', '')
            print(f'priceSquare = {priceSquare}')
            w_Sheet.write(row, 2, priceSquare)
        except:
            pass
        try:
            for q in driver.find_elements('xpath', '//*[@class="s196eif3"]'):
                if q.text.find('комнат') != -1:
                    typeRooms = q.text
                    print(f'typeRooms = {typeRooms}')
                    w_Sheet.write(row, 3, typeRooms)
                if q.text.find('м2') != -1:
                    square = q.text[:q.text.find("м2") - 1]
                    print(f'square = {square}')
                    w_Sheet.write(row, 4, square)
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
                    otdelka = q.text[q.text.find('Отделка') + 8:]
                    print(f'otdelka = {otdelka}')
                    w_Sheet.write(row, 6, otdelka)
                if q.text.find('Жилой комплекс') != -1:
                    try:
                        complexName = q.find_element('xpath', 'div[3]/a').text
                        print(f'complexName = {complexName}')
                        hrefComplex = q.find_element('xpath', 'div[3]/a').get_attribute('href')
                        print(f'hrefComplex = {hrefComplex}')
                        w_Sheet.write(row, 7, xlwt.Formula(f'HYPERLINK("{hrefComplex}"; "{complexName}")'))
                    except:
                        print("Комлекс не обнаружен")
                if q.text.find('Застройщик') != -1:
                    try:
                        zastroychikName = q.find_element('xpath', 'div[3]/a').text
                        print(f'zastroychikName = {zastroychikName}')
                        hrefZastroychik = q.find_element('xpath', 'div[3]/a').get_attribute('href')
                        print(f'hrefзастройщик = {hrefZastroychik}')
                        w_Sheet.write(row, 8, xlwt.Formula(f'HYPERLINK("{hrefZastroychik}"; "{zastroychikName}")'))
                    except:
                        print("Застройщик не найден")
                if (q.text.find('Актуально') != -1) | (q.text.find('Опубликовано') != -1):
                    actualnoNa = q.find_element('xpath', 'div[3]').text
                    print(f'actualnoNa = {actualnoNa}')
                    w_Sheet.write(row, 9, actualnoNa)
                if q.text.find('Год постройки') != -1:
                    yearOfBuilding = q.find_element('xpath', 'div[3]').text
                    print(f'yearOfBuilding = {yearOfBuilding}')
                    w_Sheet.write(row, 10, yearOfBuilding)

        except Exception as e:
            print(e)
        if hrefComplex is not None:
            complexData = complexesDict.get(complexName)
            if complexData is not None:
                w_Sheet.write(row, 11, complexData.get('startBuildingDate'))
                w_Sheet.write(row, 12, complexData.get('endBuildingDate'))
            else:
                try:
                    driver.execute_script("window.open('');")
                    tabs = driver.window_handles
                    driver.switch_to.window(tabs[2])
                    driver.get(hrefComplex)
                    buildingDates = driver.find_elements('xpath', '//*[@class="s3r6gfz"]')
                    startBuildingDate = buildingDates[0].text
                    endBuildingDate = buildingDates[1].text
                    complexesDict[complexName] = {'startBuildingDate': startBuildingDate,
                                                  'endBuildingDate': endBuildingDate}
                    print(f'startBuildingDate = {startBuildingDate}')
                    w_Sheet.write(row, 11, startBuildingDate)
                    print(f'endBuildingDate = {endBuildingDate}')
                    w_Sheet.write(row, 12, endBuildingDate)
                    driver.close()
                    driver.switch_to.window(tabs[1])
                except:
                    driver.close()
                    driver.switch_to.window(tabs[1])
        soldComplexesNum = 0
        w_Sheet.write(row, 13, soldComplexesNum)
        if hrefZastroychik is not None:
            zastroychikData = zastroychiksDict.get(zastroychikName)
            if zastroychikData is not None:
                w_Sheet.write(row, 13, zastroychikData.get('soldComplexesNum'))
            else:
                try:
                    driver.execute_script("window.open('');")
                    tabs = driver.window_handles
                    driver.switch_to.window(tabs[2])
                    driver.get(hrefZastroychik)
                    zastroychikData = driver.find_elements('xpath', '//*[@class="secfd27"]')
                    soldComplexesNum = 0
                    for sentence in zastroychikData:
                        if sentence.text.find("продано") != -1:
                            soldComplexesNum = sentence.text.split()[0]
                            break
                    zastroychiksDict[zastroychikName] = {'soldComplexesNum': soldComplexesNum}
                    print(f'soldComplexesNum = {soldComplexesNum}')
                    w_Sheet.write(row, 13, soldComplexesNum)

                    driver.close()
                    driver.switch_to.window(tabs[1])
                except:
                    driver.close()
                    driver.switch_to.window(tabs[1])
        print(f'complexesDict = {complexesDict}')
        print(f'zastroychiksDict = {zastroychiksDict}')
        w_Book.save(file_parsed_data)
        driver.close()
        driver.switch_to.window(tabs[0])
        row += 1
    w_Book.save(file_parsed_data)

driver.close()
driver.quit()
