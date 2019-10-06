import argparse
import json
import logging
import re
import sys

import requests
from openpyxl import load_workbook

from utils import open_xls_as_xlsx

API_DOMAIN_URL = "https://api.madkudu.com/v1/companies"
API_PERSON_URL = "https://api.madkudu.com/v1/persons"

logger = logging.getLogger('bulk_score')
logger.addHandler(logging.StreamHandler(sys.stdout))
logger.setLevel(logging.DEBUG)


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
    regex = re.compile('(?:@)?(?P<tld>[\w\-]+\.\w+)')

    domains_scored = {}
    emails_scored = {}

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
                print("scoring: " + person["email"])
                if not search:
                    continue

                if score_type == 'domain':
                    domain = person["email"]

                    if domain not in domains_scored:
                        params = {"domain": domain}
                        print(params)
                        resp = requests.get(API_DOMAIN_URL, auth=(api_key, ''), params=params)
                        customer_fit_result = resp.json()['properties']['customer_fit']
                        print(customer_fit_result)
                        domains_scored[domain] = customer_fit_result

                    customer_fit = domains_scored[domain]
                    new_row = "{},{},{},{}\n".format(domain, customer_fit['segment'], customer_fit['score'], customer_fit['top_signals_formatted'])
                    print(new_row)
                    result.write(new_row)

                # if score_type == 'domain':
                #     domain = person["email"]

                #     if domain not in domains_scored:
                #         params = {"email": "user@{}".format(domain)}
                #         print(params)
                #         resp = requests.get(API_PERSON_URL, auth=(api_key, ''), params=params)
                #         customer_fit_result = resp.json()['properties']['customer_fit']
                #         print(customer_fit_result)
                #         domains_scored[domain] = customer_fit_result

                #     customer_fit = domains_scored[domain]
                #     new_row = "{},{},{},{}\n".format(domain, customer_fit['segment'], customer_fit['score'], customer_fit['top_signals_formatted'])
                #     print(new_row)
                #     result.write(new_row)

                if score_type == 'email':
                    email = person["email"]
                    if email not in emails_scored:
                        params = {"email": email}
                        resp = requests.get(API_PERSON_URL, auth=(api_key, ''), params=params)
                        customer_fit_result = resp.json()['properties']['customer_fit']
                        emails_scored[email] = customer_fit_result

                    customer_fit = emails_scored[email]      
                    new_row = "{},{},{},{}\n".format(domain, customer_fit['segment'], customer_fit['score'], customer_fit['top_signals_formatted'])
                    print(new_row)
                    result.write(new_row)        
        except Exception:
            result.flush()
            logger.exception("Exception met. Relaunch to resume!\n")
            exit(1)
        exit(0)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Sends bulk persons to be scored.')
    parser.add_argument("--filename", help="xlsx file containing all the persons to score", required=True)
    parser.add_argument("--api_key", help="api key", required=True)
    parser.add_argument("--score_type", help="which score to use: either by domain or by personal email", required=True, choices=['domain', 'email'])
    parser.add_argument("--column_idx", help="domain/email column idx (i.e: BQ)", required=True)

    run_xls(**vars(parser.parse_args()))
