from openpyxl import load_workbook
import os

AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY")




def get_stocks_shares_information(file, min_row, max_row, min_col, max_col, values_only=True):
    # Get requesting stock shares information with specific row and column

    workbook = load_workbook(filename=file)
    sheet = workbook.active
    stock_shares = {}
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=values_only):
        stock_id = row[1]
        stock = {
            "Ticker": row[1],
            "Company Name": row[2],
            "Sector": row[3],
            "Industry": row[4],
            "Headquarter": row[5]
        }
        stock_shares[stock_id] = stock
    return stock_shares


if __name__ == '__main__':

    stock_info = get_stocks_shares_information(
        file="dividend_file.xlsx", min_row=7, max_row=10, min_col=2, max_col=8)
    print(stock_info)


