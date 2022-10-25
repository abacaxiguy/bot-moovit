import requests
from bs4 import BeautifulSoup
from openpyxl import *
from openpyxl.styles import Font

url = "https://moovitapp.com/index/pt-br/transporte_p%C3%BAblico-line-0903-Maceio-4466-2221160-91572044-3"
page = requests.get(url)
soup = BeautifulSoup(page.content, 'html.parser')


def get_stops_addresses():
    stops_container = soup.find_all('li', class_='stop-container')
    stops_wrappers = [stop.find('div', class_='stop-wrapper') for stop in stops_container]
    stops_addresses = []
    for stop in stops_wrappers:
        stop_address = stop.find('h3').text
        stops_addresses.append(stop_address)
    return stops_addresses


def get_line_number():
    title = soup.find('div', class_='line-image-container')
    line_number = title.find('h1', class_="text").text

    return line_number


def get_title(line_number):
    line_title = soup.find('div', class_='line-title').find('h2', class_="title").text

    return line_number + " - " + line_title


def save_into_sheet():
    sheets = Workbook()
    sheet = sheets.active

    line_number = get_line_number()
    stops = get_stops_addresses()

    # increase the size of column A
    sheet.column_dimensions['A'].width = 150

    sheet.cell(row=1, column=1).value = get_title(line_number)
    sheet.cell(row=1, column=1).font = Font(bold=True, size=20)

    for i in range(len(stops)):
        sheet.cell(row=i + 2, column=1).value = stops[i]

    sheets.save(f'linhas/{line_number}.xlsx')


save_into_sheet()