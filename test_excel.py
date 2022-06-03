import argparse
import pprint

import xlrd
import openpyxl
import pylightxl
import pandas


def parse_with_xlrd(path, sheet_index=0):
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)
    row_idx = 0
    while True:
        try:
            yield sheet.row_values(row_idx)
        except IndexError:
            break
        row_idx += 1


def parse_with_openpyxl(path, sheet_index=0):
    wb = openpyxl.load_workbook(filename=path, read_only=True)
    sheet = wb[wb.sheetnames[sheet_index]]
    for values in sheet.values:
        yield values


def parse_with_pylightxl(path, sheet_index=0):
    db = pylightxl.readxl(fn=path)
    for row in db.ws(ws=db.ws_names[sheet_index]).rows:
        yield row


def parse_with_pandas(path):
    df = pandas.read_excel(path)
    return df


def main(input_path):
    result = list(parse_with_xlrd(input_path))
    print('\n----- xlrd==1.2.0 result -----\n')
    pprint.pprint(result)

    result = list(parse_with_pylightxl(input_path))
    print('\n----- pylightxl result -----\n')
    pprint.pprint(result)

    dataframe = parse_with_pandas(input_path)
    print('\n----- pandas result -----\n')
    # pprint.pprint(dataframe)
    pprint.pprint(dataframe.dtypes)
    pprint.pprint([dataframe.columns.tolist()] + dataframe.values.tolist())

    result = list(parse_with_openpyxl(input_path))[:100]
    print('\n----- openpyxl result -----\n')
    pprint.pprint(result)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        'input_path',
    )
    args = parser.parse_args()

    main(args.input_path)
