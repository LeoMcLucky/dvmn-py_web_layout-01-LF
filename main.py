import datetime
import pandas
import collections
import argparse
from environs import Env


from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

COMPANY_BIRTH = 1920


def parse_input(default_excel_path):

    parser = argparse.ArgumentParser(
        description=(
            "Эта программа запускает сайт магазина вина 'Новое русское вино'.\n"
            "Для использования своего Excel-файла можно указать путь прямо в командной строке.\n"
            "Если путь не указан, программа берет его из .env или по умолчанию 'beverages.xlsx'.\n\n"
            "Примеры:\n"
            "python main.py\n"
            "python main.py your_path/your_file.xlsx\n"
        ),
        epilog=(
            "В .env можно задать:\n"
            "EXCEL_PATH=your_path/your_file.xlsx\n"
            "Если не задано — будет использован 'beverages.xlsx'."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        "excel_path",
        help="Путь к Excel-файлу (если не указан — из .env или 'beverages.xlsx')",
        nargs="?",
        default=default_excel_path
    )

    args = parser.parse_args()
    return args.excel_path


def get_company_age(company_birth):
    now = datetime.datetime.now()
    company_age = now.year-company_birth
    return company_age


def get_year_word(company_age):
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


def get_drinks(exel_path):
    assortment_drinks = pandas.read_excel(
        exel_path, na_values=[], keep_default_na=False).to_dict(orient='records')

    orderly_assortment_drinks = collections.defaultdict(list)
    for drink in assortment_drinks:
        category = drink['Категория']
        orderly_assortment_drinks[category].append(drink)
    return orderly_assortment_drinks


def main():

    env_config = Env()
    env_config.read_env()
    default_excel_path = env_config.str("EXCEL_PATH", "beverages.xlsx")
    exel_path = parse_input(default_excel_path)

    company_age = get_company_age(COMPANY_BIRTH)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        cap_company_age=company_age,
        cap_years_word=get_year_word(company_age),
        assortment_drinks=get_drinks(exel_path)
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
