import numpy as np
import pandas as pd
import yfinance as yf
import xlsxwriter
import math
import time

# Scrape the list of S&P500 companies from Wikipedia
table = pd.read_html('https://en.wikipedia.org/wiki/List_of_S%26P_500_companies')
stocks = table[0]
print(stocks)

#Wiki gives us a lot of crap, we just need the symbol
stocks = stocks[['Symbol']]
stocks.columns = ['Ticker']

#yfinance requires secondary assets such as BRK.B to be listed as BRK-B
stocks['Ticker'] = stocks['Ticker'].str.replace('.', '-')

my_columns = ['Ticker', 'Price','Market Capitalization', 'Number Of Shares to Buy']
final_dataframe = pd.DataFrame(columns = my_columns)

for symbol in stocks['Ticker']:
    print(f"Fetching data for {symbol}")
    tickerData = yf.Ticker(symbol)
    retries = 5  # adjustable retry total
    for i in range(retries):
        try:
            temp_df = pd.DataFrame([
                {
                    'Ticker': symbol,
                    'Price': tickerData.info['previousClose'],
                    'Market Capitalization': tickerData.info['marketCap'],
                    'Number Of Shares to Buy': 'N/A'
                }
            ])
            final_dataframe = pd.concat([final_dataframe, temp_df], ignore_index=True)
            print(f"Data fetched for {symbol}")
            break
        except KeyError:
            print(f'Data not found for {symbol}, retrying ({i+1}/{retries})...')
            time.sleep(1)  # adjustable delay
    else:
        print(f'Failed to fetch data for {symbol} after {retries} retries, skipping.')

portfolio_size = input("Enter the value of your portfolio:")

try:
    val = float(portfolio_size)
except ValueError:
    print("That's not a number! \n Try again:")
    portfolio_size = input("Enter the value of your portfolio:")

position_size = float(portfolio_size) / len(final_dataframe.index)
for i in range(0, len(final_dataframe['Ticker'])-1):
    final_dataframe.loc[i, 'Number Of Shares to Buy'] = math.floor(position_size / final_dataframe['Price'][i])

writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, sheet_name='Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_format = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_format = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

column_formats = {
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['Number of Shares to Buy', integer_format]
                    }

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

writer.close()
