import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')


def get_company_age():
    return datetime.datetime.now().year - 1920


def get_age_string(year):
    i = year % 10
    if i == 1:
        return f'{year} год'
    elif i < 5:
        return f'{year} года'
    else:
        return f'{year} лет'


wines = pandas.read_excel('wines.xlsx',
                          sheet_name='Лист1',
                          na_values='None',
                          keep_default_na=False).sort_values(
    ['Категория', 'Цена'],
    ascending=True).to_dict(
    orient='records')

wine_catalog = {}
for wine in wines:
    wine_catalog.setdefault(wine['Категория'], []).append(wine)

rendered_page = template.render(
    company_age=get_age_string(get_company_age()),
    wine_catalog=wine_catalog
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
