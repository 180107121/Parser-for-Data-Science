import pandas as pd
from bs4 import BeautifulSoup
import requests
import openpyxl
from openpyxl import Workbook

df = pd.read_excel(r'external_links.xlsx')
headers = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36', 'accept': '*/*'}

book = Workbook()
sheet = book.active
sheet.append(['link', 'house_type', 'street', 'street_name', 'house' ,'area', 'street_type', 'build_year', 'floors', 'residential_complex', 'price', 'parking', 'internet', 'security'])
sheet.title = "internal_links"

def get_html(url, params = None):
	return(requests.get(url, headers=headers, params=params))

def get_content(html, url, df, iii):
    soup = BeautifulSoup(html, 'html5lib')
    city = str(soup.find("div", class_="offer__location offer__advert-short-info").span.get_text()).split(", ")

    area = None
    street_type = None
    build_year = None
    floors = None
    residential_complex = None
    price = None
    parking = "no"
    internet = None
    security = None

    h1 = str(soup.find("h1").get_text()).split(" ")

    if "мкр." in h1 or "мкр" in h1 or "микрорайон" in h1:
        street_type = 'мкр'
    elif "проспект" in h1:
        street_type = 'проспект'
    else:
        street_type = 'улица'

    for i in city:
        if "р-н" in i:
            i = i.replace("р-н " or " р-н", "")
            area = i
            break

    table2 = soup.find_all('div', class_="offer__info-item")

    for row2 in table2:
        # print(row2.find_all('div')[0])
        if str(row2.find_all('div')[0].get_text()) == "Дом":
            for i in str(row2.find_all('div')[2].get_text()).split(", "):
                if 'г.п.' in i:
                    build_year = i.replace(' г.п.' or 'г.п. ', "")

        if str(row2.find_all('div')[0].get_text()) == "Этаж":
            for i in str(row2.find_all('div')[2].get_text()).split(" "):
                if "из" in i:
                    floors = (str(row2.find_all('div')[2].get_text()).split(" ")[2])

        if str(row2.find_all('div')[0].get_text()) == "Жилой комплекс":
            residential_complex = (str(row2.find_all('div')[2].a.get_text()))

    price = str(soup.find('div', class_="offer__price").get_text())
    price = price.replace(" 〒", "").strip()
    price = price.replace(u'\xa0', '')

    description = soup.find('div', class_="offer__parameters")
    descrip = description.find_all('dl')
    # print(descrip[2])
    for i in descrip:
        if str(i.find_all()[0].get_text()) == "Парковка":
            if str(i.find_all()[1].get_text()) != "Нет" and str(i.find_all()[1].get_text()) != "Нету" and str(
                    i.find_all()[1].get_text()) != "Отсутствует":
                parking = "yes"

        if str(i.find_all()[0].get_text()) == "Интернет":
            internet = str(i.find_all()[1].get_text())

        if str(i.find_all()[0].get_text()) == "Безопасность":
            security = str(i.find_all()[1].get_text())

    sheet.append([df['link'][iii], df['house_type'][iii], df['street'][iii], df['street_name'][iii], df['house'][iii], area, street_type, build_year, floors, residential_complex, price, parking, internet, security])





def main(url, df, i):
    html = get_html(url)
    if html.status_code == 200:
        return(get_content(html.text, url, df, i))
    else:
        print("Error")

count = 0
for i in df.index:
    try:
        print(count)
        main(df['link'][i], df, i)
        count += 1
    except:
        print("ERROR: " + str(count))
    book.save('internal_links.xlsx')


