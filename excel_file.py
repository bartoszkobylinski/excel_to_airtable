from openpyxl import load_workbook
import os
from pyairtable import Table, Api
import time

AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY")
BASE_ID = 'app5geMpic1opZ2nQ'
TABLE_ID = 'tblF32dQfRSfA3zSC'


file = "dividend_file.xlsx"

worbook = load_workbook(filename=file)

sheet = worbook.active
sector_list = []
industry_list = []
headquarter = []
for row in sheet.iter_rows(1,500,4,6, values_only=True):
    # print(f"type {type(row)} row: {row}")
    if not isinstance(row[1], str):
        pass # print(row[1])
    else:
        industry_list.append(row[1])
    sector_list.append(row[0])

    headquarter.append((row[2]))

sector_list = set(sector_list)
# print(sector_list)
another_list = []
another_industry_list = []
for row in sheet.iter_rows(501, 1500, 4, 6, values_only=True):
    if not isinstance(row[0], str):
        pass  # print(row[1])
    else:
        another_list.append(row[0])
    if not isinstance(row[1], str):
        pass
    else:
        another_industry_list.append(row[1])


# print(set(another_list))
# sector_list = set(sector_list)
# print(sector_list)

industry_list = set(industry_list)
print(industry_list)
# new_set = set(another_list) - sector_list
# new_set = set(another_list)
# print(new_set)
new_set = set(another_industry_list)
print(new_set)
c = new_set | industry_list
print(c)
# new_industry_set = set(another_industry_list) - sector_list
# print(new_industry_set)
# print(new_set)

time.sleep(200)
# print(set(industry_list))
# print(set(headquarter))


def get_stocks_shares_information(file, min_row, max_row, min_col, max_col, values_only=True):
    # Get requesting stock shares information with specific row and column

    workbook = load_workbook(filename=file)
    sheet = workbook.active
    stock_shares = {}
    for row in sheet.iter_rows(
            min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=values_only):
        print(row)
        stock_id = row[0]
        stock = {
            "Ticker": row[0],
            "Company Name": row[1],
            "Sector": row[2],
            "Industry": row[3],
            "Headquarter": row[4]
        }
        stock_shares[stock_id] = stock
    return stock_shares


api = Api(AIRTABLE_API_KEY)
table = Table(AIRTABLE_API_KEY, BASE_ID, TABLE_ID)


def update_airtable_database(stocks, airtable_api, airtable_base_id, airtable_table_id):
    for key, value in stock_info.items():
        table.create(value)


if __name__ == '__main__':

    stock_info = get_stocks_shares_information(
        file="dividend_file.xlsx", min_row=7, max_row=10, min_col=2, max_col=6)
    # print(stock_info)
    update_airtable_database(stock_info, AIRTABLE_API_KEY, BASE_ID, TABLE_ID)

