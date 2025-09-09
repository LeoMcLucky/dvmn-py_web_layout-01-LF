import datetime
import pandas
import collections

from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

COMPANY_BIRTH = datetime.datetime(year=1920, month=1, day=1, hour=1)


def get_company_age(company_birth):
    now = datetime.datetime.now()
    company_age = now.year-company_birth.year
    return company_age


def years_word(company_age):
    if 11 <= company_age % 100 <= 14:
        word = "лет"
    else:
        last_digit = company_age % 10
        if last_digit == 1:
            word = "год"
        elif 2 <= last_digit <= 4:
            word = "года"
        else:
            word = "лет"
    return word


def get_drinks(file_exel):
    assortment_drinks = pandas.read_excel(
        file_exel, na_values=[], keep_default_na=False).to_dict(orient='records')

    orderly_assortment_drinks = collections.defaultdict(list)
    for drink in assortment_drinks:
        category = drink['Категория']
        orderly_assortment_drinks[category].append(drink)
    return orderly_assortment_drinks


def main():
    file_xlsx = 'wine3.xlsx'
    company_age = get_company_age(COMPANY_BIRTH)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        cap_company_age=company_age,
        cap_years_word=years_word(company_age),
        assortment_drinks=get_drinks(file_xlsx)
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
