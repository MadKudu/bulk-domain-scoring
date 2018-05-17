import argparse
import re

from openpyxl import load_workbook


def zip_results(filename: str):
    regex = re.compile("@[\w\-]+\.\w+")

    print("Welcome to the results zipper! Wait for the xlsx to load.")
    workbook = load_workbook(filename=filename, keep_vba=False)
    sheet = workbook.active

    result_filename = filename.replace(".xlsx", ".csv")
    print("File loaded. results/{} will be zipped to it.".format(result_filename))
    with open("results/" + result_filename, "r") as result:
        result.seek(0)
        rows = sheet.max_row
        columns = sheet.max_column
        for line in range(3, rows):
            if line % 100 == 0:
                print("Zipped {0:.3f}%".format(line / (rows * 1.) * 100.))
            domain, segment, score = result.readline().split(",")
            if not regex.search(sheet.cell(line, 6).value):
                continue  # skip invalid entries, like in bulk_score
            if domain not in sheet.cell(line, 6).value:
                print("ERROR! Zipping on weird data (trying to zip {} with {})".format(sheet.cell(line, 6).value,
                                                                                       domain))
                exit(1)
            sheet.cell(line, columns + 1).value = segment
            sheet.cell(line, columns + 2).value = int(score)
    filename_with_results = filename.replace(".xlsx", "_with-results.xlsx")
    print("Now saving to {}, this might take several minutes...".format(filename_with_results))
    workbook.save(filename_with_results)
    print("You're good to go!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Sends bulk persons to be scored.')
    parser.add_argument("--filename", help="xlsx file containing all the persons to score", required=True)

    zip_results(**vars(parser.parse_args()))