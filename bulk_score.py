import argparse
import re

import requests
from openpyxl import load_workbook

API_URL = "https://api.madkudu.com/v1/companies"


def run_xls(filename: str, api_key: str):
    print("Welcome to the bulk persons searcher! Wait for the xlsx to load.")
    workbook = load_workbook(filename=filename, keep_vba=False)
    sheet = workbook.active
    regex = re.compile("@[\w\-]+\.\w+")

    domains_scored = {}

    result_filename = filename.replace(".xlsx", ".csv")
    print("File loaded. Results will be saved to results/{}.".format(result_filename))
    with open("results/" + result_filename, "a+") as result:
        result.seek(0)
        start = sum(1 for line in result)
        try:
            rows = sheet.max_row
            for line in range(start + 3, rows):
                person = {}
                if line % 100 == 0:
                    print("Currently at {}%".format(line / (rows * 1.) * 100.))
                if sheet.cell(column=20, row=line).value is not None:
                    continue
                person['email'] = sheet.cell(column=6, row=line).value
                if not person['email']:
                    continue

                """
                person['company'] = sheet.cell(column=i, row=line).value
                person['phone'] = sheet.cell(column=i + 1, row=line).value
                person['first_name'] = sheet.cell(column=i + 2, row=line).value
                person['last_name'] = sheet.cell(column=i + 3, row=line).value
                person['title'] = sheet.cell(column=i + 4, row=line).value
                person['account_id'] = sheet.cell(column=i + 6, row=line).value
                person['website'] = sheet.cell(column=i + 7, row=line).value
                person['software'] = sheet.cell(column=i + 8, row=line).value
                person['city'] = sheet.cell(column=i + 9, row=line).value
                person['state'] = sheet.cell(column=i + 10, row=line).value
                person['zip'] = sheet.cell(column=i + 11, row=line).value
                person['country'] = sheet.cell(column=i + 12, row=line).value
                person['industry'] = sheet.cell(column=i + 13, row=line).value
                person['employees'] = sheet.cell(column=i + 14, row=line).value
                person['account_owner'] = sheet.cell(column=i + 15, row=line).value
                person['avalara_customer'] = sheet.cell(column=i + 16, row=line).value
                person['avalara_opportunity'] = sheet.cell(column=i + 17, row=line).value
                """
                if not regex.search(person["email"]):
                    continue

                domain = regex.search(person["email"]).group()[1:]

                if domain not in domains_scored:
                    params = {"domain": domain}

                    resp = requests.get(API_URL, auth=(api_key, ''), params=params)

                    domains_scored[domain] = resp.json()['properties']['customer_fit']
                result.write(
                    "{},{},{}\n".format(domain, domains_scored[domain]['segment'], domains_scored[domain]['score']))
        except Exception as e:
            result.flush()
            print(e)
            print("\nException met. Relaunch to resume!\n")
            exit(1)
        exit(0)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Sends bulk persons to be scored.')
    parser.add_argument("--filename", help="xlsx file containing all the persons to score", required=True)
    parser.add_argument("--api_key", help="api key", required=True)

    run_xls(**vars(parser.parse_args()))
