from collections import OrderedDict
import xlwings as xw  # pip install xlwings
import pandas as pd  # pip install pandas
from yahoofinancials import YahooFinancials  # pip install yahoofinancials

wb = xw.Book.caller()
sheet = wb.sheets[0]
tickers = sheet.range("B2").options(expand='down').value
number_of_shares = sheet.range("C2").options(expand='down').value
# tickers = ["TD.TO", "BNS.TO"]
# number_of_shares = [10, 20]


def main():
    pull_stocks_data()
    format_data()


def pull_stocks_data():
    """
    Steps:
    1) Pull data from Yahoo FinanceCreate an empty DataFrame
    2) Calculate DRIP data
    3) Populate stock data to excel
    """
    try:
        print(f"Iterating over the following tickers: {tickers}")
        data = YahooFinancials(tickers, concurrent=True, max_workers=8)

        dripdictionary = calculate_drip_data(data)
        populate_stock_data_to_excel(data, dripdictionary)
    except:
        print("Error")


def calculate_drip_data(data):
    dripdictionary = {}
    open_prices = data.get_open_price()
    dividend_rates = data.get_dividend_rate()

    for index, stock in enumerate(tickers):
        quarterly_dividend_rate = dividend_rates.get(stock)
        open_price = open_prices.get(stock)
        # No of investment, shares for DRIP
        if quarterly_dividend_rate:
            quarterly_dividend_rate = quarterly_dividend_rate / 4
            investment_for_drip = (open_price * open_price) / quarterly_dividend_rate
            shares_for_drip = investment_for_drip / open_price
            drip = (number_of_shares[index] * quarterly_dividend_rate) / open_price
        else:
            quarterly_dividend_rate = "N/A"
            investment_for_drip = "N/A"
            shares_for_drip = "N/A"
            drip = "N/A"
        dripdictionary[stock] = {
            "quarterly_dividend_rate": quarterly_dividend_rate,
            "investment_for_drip": investment_for_drip,
            "shares_for_drip": shares_for_drip,
            "drip": drip
        }
    return dripdictionary


def populate_stock_data_to_excel(data, dripdictionary):
    columnstitlesorder = ['open_price', 'currency', 'dividend_yield', 'payout_ratio',
                          'quarterly_dividend_rate', 'quarterly_shares_for_drip',
                          'quarterly_investment_for_drip', 'quarterly_drip']
    df = pd.DataFrame()
    open_price = data.get_open_price()
    currency = data.get_currency()

    # Dividend Fields
    dividend_yield = data.get_dividend_yield()
    payout_ratio = data.get_payout_ratio()

    for index, stock in enumerate(tickers):
        # OrderDict - needed to set column order
        new_row = OrderedDict([
            ("open_price", open_price.get(stock)),
            ("currency", currency.get(stock)),
            ("dividend_yield", dividend_yield.get(stock)),
            ("payout_ratio", payout_ratio.get(stock)),

            ("quarterly_dividend_rate", dripdictionary.get(stock)["quarterly_dividend_rate"]),
            ("quarterly_shares_for_drip", dripdictionary.get(stock)["shares_for_drip"]),
            ("quarterly_investment_for_drip", dripdictionary.get(stock)["investment_for_drip"]),
            ("quarterly_drip", dripdictionary.get(stock)["drip"])
        ])
        # Append data to DataFrame
        df = df.append(new_row, ignore_index=True)

        # Populate current stock value
        populate_stock_value(index, open_price.get(stock), number_of_shares[index])
    # needed to set column order
    df = df.reindex(columns=columnstitlesorder)
    sheet.range("G1").options(index=False).value = df


def populate_stock_value(index, open_price, number_of_shares):
    try:
        sheet.range("D" + str(index + 2)).value = number_of_shares * open_price
    except:
        print(f"there was an error")


# DRIP cell RED if value N/A or less than 1, otherwise GREEN
def format_data():
    try:
        drips = sheet.range("N2").options(expand='down').value

        for index, drip in enumerate(drips):
            if drip != "N/A" and int(drip) >= 1:
                sheet.range("N" + str(index + 2)).color = (89, 255, 77)
            else:
                sheet.range("N" + str(index + 2)).color = (255, 77, 77)
    except:
        print(f"there was an error")


if __name__ == "__main__":
    xw.Book("portfolio_tracking.xlsm").set_mock_caller()
    main()
