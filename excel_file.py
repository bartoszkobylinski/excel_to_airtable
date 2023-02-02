from openpyxl import load_workbook
import os

AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]

workbook = load_workbook(filename="dividend_file.xlsx")

sheet = workbook.active

stock_shares = {}
for row in sheet.iter_rows(min_row=7, max_row=12, min_col=2, max_col=20, values_only=True):
    stock_id = row[1]
    stock = {
        "Ticker": row[1],
        "Company Name": row[2],
        "Sector": row[3],
        "Industry": row[4],
        "Headquarter": row[5]
    }
    stock_shares[stock_id] = stock

print(stock_shares.keys())








