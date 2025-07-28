import yfinance as yf
import pandas as pd

tickers = ['AAPL', 'MSFT', 'TSLA']
data = []

for ticker in tickers:
    stock = yf.Ticker(ticker)
    info = stock.info
    pe = info.get("trailingPE", "N/A")
    roe = info.get("returnOnEquity", "N/A")
    growth = info.get("revenueGrowth", "N/A")
    data.append([ticker, pe, roe, growth])

df = pd.DataFrame(data, columns=['Ticker', 'P/E Ratio', 'ROE', 'Revenue Growth'])
df.to_excel('data/financial_data.xlsx', index=False)
print("✅ Excel file created at data/financial_data.xlsx")
