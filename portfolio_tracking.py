import time, os, sys
import xlwings as xw  # pip install xlwings
import pandas as pd  # pip install pandas
from yahoofinancials import YahooFinancials  # pip install yahoofinancials

wb = xw.Book.caller()    
sheet = wb.sheets[0]
tickers = sheet.range("B2").options(expand='down').value
number_of_shares = sheet.range("C2").options(expand='down').value

def main():
    pull_stock_data()
    format_data()

if __name__ == "__main__":
    xw.Book("portfolio_tracking.xlsm").set_mock_caller()
    main()
    
def pull_stock_data():
    """
    Steps:
    1) Create an empty DataFrame
    2) Iterate over tickers, pull data from Yahoo Finance & add data to dictonary "new row"
    3) Append "new row" to DataFrame
    4) Return DataFrame
    """
    if tickers:
        try:
            print(f"Iterating over the following tickers: {tickers}")
            df = pd.DataFrame()

            for index, ticker in enumerate(tickers):
                print(f"~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                print(f"Pulling financial data for: {ticker} ...")
                data = YahooFinancials(ticker)
                print(f"Data -> {data}")
                open_price = data.get_open_price()
                currency = data.get_currency()
                yearly_high = data.get_yearly_high()

                # Dividend Fields
                dividend_yield = data.get_dividend_yield()
                quarterly_dividend_rate = data.get_dividend_rate()
                payout_ratio = data.get_payout_ratio()

                # investment_for_drip = "N/A"
                # shares_for_drip = "N/A"             
                # drip = "N/A"
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

                # Create dictionary
                new_row = {
                    # "ticker": ticker,
                    "open_price": open_price,
                    "currency": currency,
                    # "yearly_high": yearly_high,
                    "dividend_yield": dividend_yield,
                    "quarterly_dividend_rate": quarterly_dividend_rate,
                    "payout_ratio": payout_ratio,

                    # "dividend_payment": , 
                    "quarterly_shares_for_drip": shares_for_drip,
                    "quarterly_investment_for_drip": investment_for_drip,
                    "quarterly_drip": drip
                }
                # Append data to DataFrame
                df = df.append(new_row, ignore_index=True)
                # Populate current stock value
                populate_stock_value(index, open_price, number_of_shares[index])
                
        except:
            print(f"there was an error")    

        sheet.range("G1").options(index=False).value = df              
        return df
    return pd.DataFrame()

# DRIP cell RED if value N/A or less than 1, otherwise GREEN
def format_data():
    try:
        drips = sheet.range("L2").options(expand='down').value
    
        for index, drip in enumerate(drips):
            if drip != "N/A" and int(drip) >= 1:
                sheet.range("L" + str(index + 2)).color = (89, 255, 77)
            else:
                sheet.range("L" + str(index + 2)).color = (255, 77, 77)
    except:    
        print(f"there was an error")    
        
def populate_stock_value(index, open_price, number_of_shares):
    try:
        sheet.range("D" + str(index + 2)).value = number_of_shares * open_price
    except:
        print(f"there was an error")    
