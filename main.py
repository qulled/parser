from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import random
import time
import re

s = Service("chromedriver.exe")
driver = webdriver.Chrome(service=s)
driver.maximize_window()

data = pd.read_excel(r"data.xlsx")
data = data.drop(index=[0], axis=0)
data['Артикул ВБ'] = data['Артикул ВБ'].astype(int).astype(str)

place = 0

def parcer_card(url, driver, index, article):
    try:
        driver.get(url=url)
        time.sleep(random.randrange(3, 4))
        find_raiting = driver.find_element(By.CLASS_NAME, "same-part-kt__count-review")
        find_raiting.click()
        time.sleep(random.randrange(2, 3))

        # создаем данные для парсинга цены, рейтинга и количества отзывов через BeautifulSoup
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        time.sleep(random.randrange(2, 4))

        # парсинг цены
        try:
            for i in soup.find_all('span', class_="price-block__final-price"):
                price = i.get_text().replace("₽", '')
                price = int(re.sub(r"\s+", "", price))
                data.loc[index, 'Цена на ВБ'] = price
        except Exception:
            for i in soup.find_all('div', class_="same-part-kt__sold-out-product"):
                sold_out = i.get_text().strip()
                data.loc[index, 'Цена на ВБ'] = 'Нет в наличии'
                # парсинг отсутствующих размеров (для одежды)
            # if soup.find_all('label', class_="j-size disabled"):
            # for i in soup.find_all('label', class_="j-size disabled"):
            # size = i.get_text().strip()
            # print (size)

        # парсинг рейтинга
        if soup.find_all('span', class_="user-scores__score"):
            for i in soup.find_all('span', class_="user-scores__score"):
                raiting = float(i.get_text().strip())
                data.loc[index, 'Рейтинг'] = raiting
        else:
            data.loc[index, 'Рейтинг'] = 'Нет рейтинга'

        # парсинг отзывов
        if soup.find_all('div', class_="user-scores__text-wrap"):
            for i in soup.find_all('div', class_="user-scores__text-wrap"):
                reviews = i.get_text().strip()
                reviews = re.sub(r"\s+", "", reviews)
                reviews = [int(a) for a in re.findall(r'-?\d+\.?\d*', reviews)]
                data.loc[index, 'Количество отзывов'] = reviews[0]
        else:
            data.loc[index, 'Количество отзывов'] = 'Нет отзывов'

    except Exception as _ex:
        print(_ex)
    finally:
        time.sleep(random.randrange(1, 2))
        find_retrieval(article, driver, index)


# поиск места по поисковому запросу
def find_retrieval(article, driver, index):
    try:
        # переходим в поле поиска и вводим поисковый запрос
        time.sleep(random.randrange(1, 3))
        search = driver.find_element(By.ID, "searchInput")
        search.click()
        # подтягиваем поисковый запрос из DF
        retrieval = data.iloc[index - 1]['Поисковый Запрос']
        search.send_keys(f'{retrieval}' + Keys.RETURN)
        # driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    except Exception as _ex:
        print(_ex)

    finally:
        time.sleep(random.randrange(3, 4))
        parsing_catalog(article, driver, index)

# парсинг карточек товаров в каталоге
def parsing_catalog(article, driver, index):
    container = []
    try:
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        time.sleep(random.randrange(3, 4))
        # пробегаемся по странице с карточками товаров
        for i in soup.find_all('div',
                               class_="product-card j-card-item j-advert-card-item advert-card-item "
                                      "j-good-for-listing-event"):
            # сразу отфильтровываем рекламные карточки
            if i.find('p', class_='product-card__tip-promo'):
                continue
            else:
                take = i.get("data-popup-nm-id")
                container.append(take)
        for i in soup.find_all('div', class_="product-card j-card-item j-good-for-listing-event"):
            take = i.get("data-popup-nm-id")
            container.append(take)
        for i in soup.find_all('div', class_="product-card j-card-item"):
            take = i.get("data-popup-nm-id")
            container.append(take)
    except Exception as _ex:
        print(_ex)

    finally:
        time.sleep(random.randrange(1, 2))
        search_id(article, driver, container, index)


# ищем совпадение по набору id карточек
def search_id(article, driver, container, index):
    for i in container:
        if i != article:
            global place
            place += 1
            try:
                if i == container[-1] and place < 1000:
                    next_page = driver.find_element(By.LINK_TEXT, 'Следующая страница')
                    time.sleep(random.randrange(1, 3))
                    next_page.click()
                    parsing_catalog(article, driver, index)
                elif place >= 1000:
                    data.loc[index, 'Место по поисковому запросу'] = 'Нет в рейтинге'
                    place_mod()
                    break
            except Exception as _ex:
                print(_ex)

            finally:
                data.loc[index, 'Место по поисковому запросу'] = 'Нет в рейтинге'
        else:
            data.loc[index, 'Место по поисковому запросу'] = place
            place_mod()
            break


def place_mod():
    global place
    place = 0


# тело парсера
def main(driver):
    index = 1
    count = 0
    try:
        for article in data['Артикул ВБ']:
            parcer_card(f"https://www.wildberries.ru/catalog/{article}/detail.aspx", driver, index, article)
            data.to_excel('NEW DATA.xlsx')
            index += 1
            #каждые 10 итераций доп. пауза
            if count == 10:
                count = 0
                time.sleep(random.randrange(8, 12))
    finally:
        driver.close()
        driver.quit()


if __name__ == "__main__":
    main(driver)
