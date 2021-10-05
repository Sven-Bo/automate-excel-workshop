import xlwings as xw
import yfinance as yf
import mplfinance as mpf
import matplotlib.pyplot as plt
from pathlib import Path

curr_file = Path(__file__).parent


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    ticker = sheet.range("ticker_symbol").value
    data = yf.download(ticker, start="2021-08-01", end="2021-09-1")
    output_figure = curr_file / 'mplfiance.png'
    fig = mpf.plot(data, type='candle',mav=(3,6,9),volume=True, show_nontrading=True, savefig=output_figure)
    sheet.pictures.add(output_figure, name='MyPlot', update=True)



if __name__ == "__main__":
    xw.Book("financecharts.xlsm").set_mock_caller()
    main()
