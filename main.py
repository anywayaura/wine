import argparse
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():

    COMPANY_FOUNDATION_YEAR = 1920

    parser = argparse.ArgumentParser(description='Программа запускает сайт, с базой в виде excel файла')
    parser.add_argument('catalog_file', help='excel файл каталог')
    args = parser.parse_args()


    wines_catalog_file = args.catalog_file

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')


    def get_age_string():
        years = datetime.datetime.now().year - COMPANY_FOUNDATION_YEAR
        i = years % 10
        if i == 1:
            return f'{years} год'
        elif i < 5:
            return f'{years} года'
        else:
            return f'{years} лет'

    wines = pandas.read_excel(wines_catalog_file,
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
        company_age=get_age_string(),
        wine_catalog=wine_catalog
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
