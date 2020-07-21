from bs4 import BeautifulSoup
import requests
import openpyxl
from openpyxl import Workbook

headers = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36', 'accept': '*/*'}

def get_html(url, params = None):
	return(requests.get(url, headers=headers, params=params))

#book = openpyxl.load_workbook('external_links.xlsx', data_only=True)
book = Workbook()
sheet = book.active
sheet.append(['link', 'house_type', 'street', 'street_name', 'house'])
sheet.title = "external_links"


def get_content(html, url):
    soup = BeautifulSoup(html, 'html5lib')
    table = soup.find_all('div', class_ = "a-card__inc")
    for row in table:
        #sheet.append(["https://krisha.kz/" + str(row.a.get('href'))])
        house_type, house, street_name, street = external(row, url)


        sheet.append(["https://krisha.kz/"+str(row.a.get('href')), house_type, street, street_name, house])
        #sheet.append(["https://krisha.kz/"+str(row.a.get('href'))])
        #print(str(row.find('div', class_="a-card__subtitle").get_text()).strip(), ",,,,", street_name.strip())

    book.save('external_links.xlsx')

def external(row, url):
    house_type = None
    house = None
    street_name = None

    street = str(row.find('div', class_="a-card__subtitle").get_text()).strip().split(", ")[-1]
    street_name = street
    for i in reversed(street.split(" ")):
        if i[0:1].isdigit() or i[-1:-2].isdigit():
            house = i
            break


    try:
        if house in street:
            street_name = street.replace(" " + str(house), "")
    except TypeError:
        pass
    if "kvartiry" in str(url):
        house_type = "квартиры"
    elif "doma" in str(url):
        house_type = "частный дом"

    return house_type, house, street_name, street

def main(url):
    html = get_html(url)
    if html.status_code == 200:
        return(get_content(html.text, url))
    else:
        print("Error")

for i in range(1,1001,1):
    try:
        main("https://krisha.kz/prodazha/kvartiry/?page=" + str(i))
        print(i)
    except:
        print ("ERROR: " + str(i))
        i = i - 1
