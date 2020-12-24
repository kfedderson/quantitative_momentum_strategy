import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
import os
from scipy.stats import percentileofscore as pos
from statistics import mean


SB_TOKEN = os.getenv('SB_TOKEN')
stocks = pd.read_csv('sp_500_stocks.csv')
my_columns = [
  'Ticker', 
  'Stock Price', 
  'One Year Return',
  'One Year Percentile',
  'Six Month Return',
  'Six Month Percentile',
  'Three Month Return',
  'Three Month Percentile',
  'One Month Return',
  'One Month Percentile',
  'HQM Score',
  'Shares to buy'
  ]

PORTFOLIO_SIZE = 1000000.00

def chunks(lst, n):
  for i in range(0,len(lst), n):
    yield lst[i:i + n]

# Add stocks to pandas dataframe, batch API call
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
  symbol_strings.append(','.join(symbol_groups[i]))

final_dataframe = pd.DataFrame(columns = my_columns)

for symbol_string in symbol_strings:
  api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=price,stats&token={SB_TOKEN}'
  data = requests.get(api_url).json()
  for symbol in symbol_string.split(','):
    final_dataframe = final_dataframe.append(
      pd.Series(
        [
          symbol,
          data[symbol]['price'],
          data[symbol]['stats']['year1ChangePercent'],
          'N/A',
          data[symbol]['stats']['month6ChangePercent'],
          'N/A',
          data[symbol]['stats']['month3ChangePercent'],
          'N/A',
          data[symbol]['stats']['month1ChangePercent'],
          'N/A',
          'N/A',
          'N/A'
        ],
        index = my_columns
      ),
      ignore_index = True
    )

final_dataframe.fillna(value = 0, inplace = True)

time_periods = [
  'One Year',
  'Six Month',
  'Three Month',
  'One Month'
  ]

for row in final_dataframe.index:
  momentum_percentiles = []
  
  for time_period in time_periods:
    change_col = f'{time_period} Return'
    percent_col = f'{time_period} Percentile'

    final_dataframe.loc[row, percent_col] = pos(final_dataframe[change_col], final_dataframe.loc[row, change_col])

    momentum_percentiles.append(final_dataframe.loc[row, percent_col])
    final_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)


# Sort by HQM Score
final_dataframe.sort_values('HQM Score', ascending = False, inplace = True)

# Keep highest 50
final_dataframe = final_dataframe[:50]
final_dataframe.reset_index(inplace = True, drop = True)

position_size = PORTFOLIO_SIZE/len(final_dataframe.index)

# Calculate number of shares to buy for each stock and write to dataframe
for i in range(0, len(final_dataframe.index)):
  final_dataframe.loc[i, 'Shares to buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

# Write results out to excel file with formatting
writer = pd.ExcelWriter('recommended_trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

bg_color = '#0a0a23'
ft_color = '#ffffff'

string_format = writer.book.add_format(
  {
    'font_color': ft_color,
    'bg_color': bg_color,
    'border': 1
  }
)

dollar_format = writer.book.add_format(
  {
    'num_format': '$0.00',
    'font_color': ft_color,
    'bg_color': bg_color,
    'border': 1
  }
)

int_format = writer.book.add_format(
  {
    'num_format' : '0',
    'font_color' : ft_color,
    'bg_color' : bg_color,
    'border' : 1
  }
)

pct_format = writer.book.add_format(
  {
    'num_format' : '0.0%',
    'font_color' : ft_color,
    'bg_color' : bg_color,
    'border' : 1
  }
)

column_formats = {
  'A': ['Ticker', string_format],
  'B': ['Stock Price', dollar_format],
  'C': ['One Year Return', dollar_format],
  'D': ['One Year Percentile', pct_format],
  'E': ['Six Month Return', dollar_format],
  'F': ['Six Month Percentile', pct_format],
  'G': ['Three Month Return', dollar_format],
  'H': ['Three Month Percentile', pct_format],
  'I': ['One Month Return', dollar_format],
  'J': ['One Month Percentile', pct_format],
  'K': ['HQM Score', pct_format],
  'L': ['Shares to buy', int_format]
}

for column in column_formats.keys():
  writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
  writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

writer.save()