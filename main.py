import argparse
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_age_string(foundation):
    years = datetime.datetime.now().year - foundation
    i = years % 10
    if i == 1:
        return f'{years} год'
    elif i < 5:
        return f'{years} года'
    else:
        return f'{years} лет'


def main():

    company_foundation_year = 1920

    parser = argparse.ArgumentParser(description='Программа запускает сайт, с базой в виде excel файла')
    parser.add_argument('catalog_filepath', help='путь к excel каталогу')
    args = parser.parse_args()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    wines = pandas.read_excel(args.catalog_filepath,
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
        company_age=get_age_string(company_foundation_year),
        wine_catalog=wine_catalog
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
