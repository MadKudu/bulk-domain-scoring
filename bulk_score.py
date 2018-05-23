import sys

import xlrd
from openpyxl.workbook import Workbook
import argparse
import re
import logging

import requests
from openpyxl import load_workbook

API_DOMAIN_URL = "https://api.madkudu.com/v1/companies"
API_PERSON_URL = "https://api.madkudu.com/v1/companies"

logger = logging.getLogger('bulk_score')
logger.addHandler(logging.StreamHandler(sys.stdout))
logger.setLevel(logging.DEBUG)


def open_xls_as_xlsx(filename):
    # first open using xlrd
    book = xlrd.open_workbook(filename)
    index = 0
    nrows, ncols = 0, 0
    while nrows * ncols == 0:
        sheet = book.sheet_by_index(index)
        nrows = sheet.nrows
        ncols = sheet.ncols
        index += 1

    # prepare a xlsx sheet
    book1 = Workbook()
    sheet1 = book1.active

    for row in range(0, nrows):
        for col in range(0, ncols):
            sheet1.cell(row=row+1, column=col+1).value = sheet.cell_value(row, col)
    return book1


def run_xls(filename: str, api_key: str, score_type: str, column_idx: int):
    print("Welcome to the bulk persons searcher! Wait for the xlsx to load.")
    if re.search('\.xlsx$', filename):
        workbook = load_workbook(filename=filename, keep_vba=False)
        result_filename = filename.replace(".xlsx", ".csv")
    elif re.search('\.xls$', filename):
        workbook = open_xls_as_xlsx(filename)
        result_filename = filename.replace(".xls", ".csv")
    else:
        print("Unsupported file format!")
        exit(1)

    sheet = workbook.active
    regex = re.compile(r'(?:@)?(?P<tld>[\w\-]+\.\w+)')

    domains_scored = {}

    print("File loaded. Results will be saved to results/{}.".format(result_filename))
    with open("results/" + result_filename, "a+") as result:
        result.seek(0)
        start = sum(1 for line in result)
        skip_empty = 2
        for row in sheet['A2:B256']:
            if not row[0].value:
                skip_empty += 1
            else:
                break
        try:
            rows = sheet.max_row
            for line in range(start + skip_empty, rows):
                person = {}
                if line % 100 == 0:
                    print("Currently at {}%".format(line / (rows * 1.) * 100.))
                person['email'] = sheet['{}{}'.format(column_idx, line)].value
                if not person['email']:
                    continue

                search = regex.search(person["email"])
                if not search:
                    continue

                domain = search.group('tld')

                if domain not in domains_scored:
                    params = {"domain": domain}

                    resp = requests.get(API_DOMAIN_URL, auth=(api_key, ''), params=params)

                    domains_scored[domain] = resp.json()['properties']['customer_fit']
                result.write(
                    "{},{},{}\n".format(domain, domains_scored[domain]['segment'], domains_scored[domain]['score']))
        except Exception:
            result.flush()
            logger.exception("Exception met. Relaunch to resume!\n")
            exit(1)
        exit(0)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Sends bulk persons to be scored.')
    parser.add_argument("--filename", help="xlsx file containing all the persons to score", required=True)
    parser.add_argument("--api_key", help="api key", required=True)
    parser.add_argument("--score_type", help="which score to use: either by domain or by personal email", required=True, choices=['domain', 'mail'])
    parser.add_argument("--column_idx", help="domain/mail column idx (i.e: BQ)", required=True)

    run_xls(**vars(parser.parse_args()))
