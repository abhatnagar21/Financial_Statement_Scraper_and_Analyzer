from flask import Flask, render_template, request, send_file
import yfinance as yf
import pandas as pd
import os

app = Flask(__name__)

def extract_financials(ticker):
    stock = yf.Ticker(ticker)
    balance_sheet = stock.balance_sheet
    income_statement = stock.financials
    cashflow_statement = stock.cashflow

    financials_dict = {
        "Balance Sheet": balance_sheet,
        "Income Statement": income_statement,
        "Cashflow Statement": cashflow_statement
    }
    return financials_dict

def save_to_excel(financials_dict, filename="financials.xlsx"):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for sheet_name, df in financials_dict.items():
            if not df.empty:
                df.to_excel(writer, sheet_name=sheet_name)
    return filename

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        ticker = request.form['ticker'].upper()
        financials = extract_financials(ticker)
        filepath = save_to_excel(financials)
        return send_file(filepath, as_attachment=True)
    return render_template('index_fs.html')

if __name__ == "__main__":
    app.run(debug=True)
