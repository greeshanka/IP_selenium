import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import openpyxl


def get_data_with_selenium(url):
    try:
        path = Service("C:/projects/_IP_testovoe/pars/chromedriver.exe")
        driver = webdriver.Chrome(service=path)
        driver.get(url=url)
        time.sleep(6)

        with open("catalog_selenium.html", "w", encoding="utf-8") as file:
            file.write(driver.page_source)

    except Exception as ex:
        print(ex)
    finally:
        driver.close()
        driver.quit()


def read_catalog_and_insert_in_exel_file():
    with open("catalog_selenium.html", encoding="utf-8") as file:
        scr = file.read()

    wb = openpyxl.Workbook()
    wb.create_sheet(title='Первый лист', index=0)
    sheet = wb['Первый лист']
    count = 2
    soup = BeautifulSoup(scr, 'lxml')

    sheet['A1'] = 'Ссылка'
    sheet['B1'] = 'Наименование товара'
    sheet['C1'] = 'Ценник'
    sheet['D1'] = 'Рейтинг'

    list_catalog = []

    list__products_urls = soup.find_all('div', class_="col-mbs-12 col-mbm-6 col-xs-4 col-md-3")
    for url in list__products_urls:
        url_exel = "https://kazanexpress.ru" + url.find('a').get("href")
        sheet = wb['Первый лист']
        sheet[f'A{count}'] = url_exel
        count += 1
        list_catalog.append(url_exel)

    count = 2

    for title in list__products_urls:
        title_exel = title.find('a').get('title')
        sheet = wb['Первый лист']
        sheet[f'B{count}'] = title_exel
        count += 1

    count = 2

    list__products_prices = soup.find_all('span', class_='currency product-card-price slightly medium')
    for i in list__products_prices:
        price_exel = i.text
        sheet[f'C{count}'] = price_exel
        count += 1

    count = 2

    list__products_ratings = soup.find_all('span', class_='orders')
    for rating in list__products_ratings:
        rating_exel = rating.text
        sheet[f'D{count}'] = rating_exel
        count += 1
    wb.save('catalog.xlsx')
    print(list_catalog)


def main():
    get_data_with_selenium("https://kazanexpress.ru/search?query=%D1%84%D1%83%D1%82%D0%B1%D0%BE%D0%BB%D0%BA%D0%B0")
    read_catalog_and_insert_in_exel_file()


if __name__ == '__main__':
    main()
