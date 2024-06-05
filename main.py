import time
import random
import urllib.request
from openpyxl import Workbook
from selenium import webdriver
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image
from selenium.webdriver.common.by import By


def run_web_driver() -> webdriver.Chrome:
    """
    Функция инициализации драйвера Chrome
    Возвращаемое значение: экземпляр класса webdriver.Chrome
    """
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")
    #Создаем объект драйвера

    driver = webdriver.Chrome(options=options)
    return driver


def get_links(driver: webdriver.Chrome, pages: int = 5) -> list:
    """
    Функция для получения ссылок для каждого фильма
    Описание аргументов: driver - веб-драйвер хрома,
    pages - количество страниц, на которых необходимо парсить ссылки
    Возвращаемое значение: список ссылок на фильмы
    """
    links = []
    for page in range(1, pages+1):
        #Получение новой страницы с фильмами
        driver.get(f'https://www.kinopoisk.ru/lists/movies/top250/?page={page}')
        if page == 1:
            #Ожидание для капчи
            time.sleep(10)
        else:
            time.sleep(int(random.random()*5))
        #Получение всех контейнеров с фильмами на странице
        container_elements = driver.find_elements(By.CLASS_NAME, 'styles_root__ti07r')
        #Получение списка фильмов на странице
        for element in container_elements:
            a = element.find_elements(By.TAG_NAME, 'a')
            links.append(a[0].get_attribute("href"))
    return links


def get_film_info(driver: webdriver.Chrome, links: list) -> list:
    """
    Функция для получения информации по фильмам
    Описание аргументов: driver - веб-драйвер хрома,
    links - список ссылок на фильмы, которые необходи распарсить
    Возвращаемое значение: Список, который содержит списки с информацией о фильмах
    """
    listFilmInfo = []
    imgCnt = 0
    #Парсинг данных
    for link in links:
        #Список для хранения данных по фильму
        filmInfo = []
        #Переход на страницу с фильмом
        driver.get(link)
        #Получения картинки preview
        preview_img = driver.find_element(By.CLASS_NAME, 'film-poster')
        image_url = preview_img.get_attribute('src')
        urllib.request.urlretrieve(image_url, f'img/{imgCnt}.png')
        filmInfo.append(f'img/{imgCnt}.png')
        imgCnt += 1
        #preview = driver.find_element(By.CLASS_NAME, 'styles_root__aZJRN')
        #filmInfo.append(preview.text)
        # Получения текста рейтинга
        rating = driver.find_element(By.CLASS_NAME, 'styles_ratingKpTop__84afd')
        filmInfo.append(rating.text)
        #Получения текста названия
        name = driver.find_element(By.TAG_NAME, 'h1')
        filmInfo.append(name.text.split('(')[0])
        #Получения текста описания
        desc = driver.find_element(By.CLASS_NAME, 'styles_paragraph__wEGPz')
        filmInfo.append(desc.text)
        # Получения года выпуска фильма
        try:
            year = driver.find_element(By.CLASS_NAME, 'styles_linkLight__cha3C')
        except:
            year = driver.find_element(By.CLASS_NAME, 'styles_linkDark__7m929')
        filmInfo.append(int(year.text))

        listFilmInfo.append(filmInfo)
        print(filmInfo)
        time.sleep(int(random.random() * 5))
    return listFilmInfo




def sort_film_list(listFilmInfo: list):
    """
    Функция для сортировки списков с информацией о фильмах по году
    Описание аргументов: listFilmInfo - Список, который содержит списки с информацией о фильмах
    """
    listFilmInfo.sort(key=lambda x: x[-1], reverse=True)


def put_by_excel(listFilmInfo: list, filename: str):
    """
    Функция записи информации о фильмах в Excel таблицу
    Описание аргументов: listFilmInfo - Список, который содержит списки с информацией о фильмах
    filename - имя Excel файла
    """
    #Создание экземпляра для работы с таблицами Excel
    workbook = Workbook()
    sheet = workbook.active

    #Формируем строку оглавления и настраиваем ширину колонок
    sheet['A1'] = "Preview"
    sheet['B1'] = "Rating"
    sheet['C1'] = "Name"
    sheet['D1'] = "Description"
    sheet['E1'] = "Year"
    sheet.column_dimensions['A'].width = 14.1
    sheet.column_dimensions['B'].width = 7
    sheet.column_dimensions['C'].width = 16
    sheet.column_dimensions['D'].width = 80
    sheet.column_dimensions['E'].width = 7
    row_num = 2
    for film in listFilmInfo:
        #Настраиваем высоту строки
        sheet.row_dimensions[row_num].height = 112.5
        for col_num, value in enumerate(film, start=1):
            #Добавление картинки
            if col_num == 1:
                img = Image(value)
                img.height = 150
                img.width = 100
                sheet.add_image(img,f'A{row_num}')
            #Добавление остальной информации
            sheet.cell(row=row_num, column=col_num, value=value).alignment = Alignment(wrap_text=True)
        row_num += 1

    workbook.save(filename)

if __name__ == '__main__':
    driver_ = run_web_driver()
    links_  = get_links(driver_)
    listFilmInfo_ = get_film_info(driver_, links_)
    sort_film_list(listFilmInfo_)
    put_by_excel(listFilmInfo_, 'FilmList1.xlsx')

