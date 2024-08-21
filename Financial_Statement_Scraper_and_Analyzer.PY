import yfinance as yf
import pandas as pd

def extract_financials(ticker):
    stock = yf.Ticker(ticker)
    
    # Extract financial statements
    balance_sheet = stock.balance_sheet
    income_statement = stock.financials
    cashflow_statement = stock.cashflow

    # Debug: Print the first few rows of each DataFrame
    print("Balance Sheet data:\n", balance_sheet.head())
    print("Income Statement data:\n", income_statement.head())
    print("Cashflow Statement data:\n", cashflow_statement.head())
    
    # Check if the data is empty
    if balance_sheet.empty:
        print("Balance Sheet data is empty.")
    if income_statement.empty:
        print("Income Statement data is empty.")
    if cashflow_statement.empty:
        print("Cashflow Statement data is empty.")
    
    financials_dict = {
        "Balance Sheet": balance_sheet,
        "Income Statement": income_statement,
        "Cashflow Statement": cashflow_statement
    }
    
    return financials_dict

def save_to_excel(financials_dict, filename="financials.xlsx"):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for sheet_name, df in financials_dict.items():
            # Debug: Print the first few rows of each DataFrame
            print(f"{sheet_name} data:\n", df.head())
            
            if not df.empty:
                df.to_excel(writer, sheet_name=sheet_name)
            else:
                print(f"No data to save for {sheet_name}")

    print(f"Financials saved to {filename}")

def analyze_financials(financials_dict):
    # Extract data from financials_dict
    balance_sheet = financials_dict.get("Balance Sheet")
    income_statement = financials_dict.get("Income Statement")
    
    # Debug: Print available row labels
    print("Available Balance Sheet row labels:\n", balance_sheet.index.tolist())
    print("Available Income Statement row labels:\n", income_statement.index.tolist())
    
    if not balance_sheet.empty and not income_statement.empty:
        # Handle missing data with a fallback approach
        try:
            total_assets = balance_sheet.loc["Total Assets"].iloc[0]
        except KeyError:
            total_assets = None
            print("Total Assets not found in Balance Sheet.")
        
        try:
            total_liabilities = balance_sheet.loc["Total Liabilities Net Minority Interest"].iloc[0]
        except KeyError:
            total_liabilities = None
            print("Total Liabilities not found in Balance Sheet.")
        
        try:
            equity = balance_sheet.loc["Total Equity Gross Minority Interest"].iloc[0]
        except KeyError:
            equity = None
            print("Total Equity not found in Balance Sheet.")
        
        try:
            net_income = income_statement.loc["Net Income"].iloc[0]
        except KeyError:
            net_income = None
            print("Net Income not found in Income Statement.")
        
        try:
            revenue = income_statement.loc["Total Revenue"].iloc[0]
        except KeyError:
            revenue = None
            print("Total Revenue not found in Income Statement.")
        
        # Calculate financial ratios only if all necessary data is available
        if all(v is not None for v in [total_assets, total_liabilities, equity, net_income, revenue]):
            try:
                current_assets = balance_sheet.loc["Total Current Assets"].iloc[0]
                current_liabilities = balance_sheet.loc["Total Current Liabilities"].iloc[0]
                current_ratio = current_assets / current_liabilities
            except KeyError:
                current_ratio = None
                print("Current Ratio data not found.")
            
            debt_to_equity_ratio = total_liabilities / equity if equity else None
            return_on_assets = net_income / total_assets if total_assets else None
            return_on_equity = net_income / equity if equity else None
            profit_margin = net_income / revenue if revenue else None
            
            # Print financial metrics
            print("\nFinancial Analysis:")
            if current_ratio is not None:
                print(f"Current Ratio: {current_ratio:.2f}")
            if debt_to_equity_ratio is not None:
                print(f"Debt to Equity Ratio: {debt_to_equity_ratio:.2f}")
            if return_on_assets is not None:
                print(f"Return on Assets (ROA): {return_on_assets:.2f}")
            if return_on_equity is not None:
                print(f"Return on Equity (ROE): {return_on_equity:.2f}")
            if profit_margin is not None:
                print(f"Profit Margin: {profit_margin:.2f}")
            
            # Decision based on financial metrics
            if current_ratio is not None and current_ratio > 1.5 and \
               debt_to_equity_ratio is not None and debt_to_equity_ratio < 1 and \
               return_on_assets is not None and return_on_assets > 0.05 and \
               return_on_equity is not None and return_on_equity > 0.1 and \
               profit_margin is not None and profit_margin > 0.1:
                print("The company is financially healthy and might be worth investing in.")
            else:
                print("The company may not be financially healthy. Further analysis is recommended.")
        else:
            print("Insufficient financial data for analysis.")
    else:
        print("Insufficient financial data for analysis.")

# Example usage
ticker = input("Enter the ticker symbol: ").upper()
financials = extract_financials(ticker)
save_to_excel(financials)
analyze_financials(financials)
