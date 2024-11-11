import pandas as pd

#load excel file
file_path = 'Stock Valuation Dividend Discount Model.xlsx'
excel_file = pd.ExcelFile(file_path)

dividends_hist = excel_file.parse('Dividends History')
coe = excel_file.parse('CoE')
fair_share_price_calc = excel_file.parse('Fair Share Price Calc')

# Clean Dividends History sheet
dividends_hist.columns = [
    "Year", "Payment Date", "Record Date", "Amount", "Value1", "Value2",
    "Year_Indicator", "Yearly Dividend", "Dividend Growth One Year", "Med Growth"
]
dividends_hist = dividends_hist[dividends_hist["Year"].notna() | dividends_hist["Payment Date"].notna()]

# Calculate growth rates
dividends_hist["Yearly Dividend"] = pd.to_numeric(dividends_hist["Yearly Dividend"], errors="coerce")
dividends_hist["Calculated Growth"] = dividends_hist["Yearly Dividend"].pct_change(fill_method=None)

# Extract values for CoE calculation
risk_free_rate = 0.043  # 4.30% 
market_return = 0.0637  # 6.37%
beta = 0.59

# Calculate CoE
calculated_coe = risk_free_rate + beta * (market_return - risk_free_rate)

# Clean up Fair Share Price Calc 
fair_share_price_calc.columns = ["Metric", "Value1", "Value2", "Unused"]
## print(fair_share_price_calc["Metric"].unique())


# Extract values for the Dividend Discount Model (DDM) calculation
## last_dividend = pd.to_numeric(fair_share_price_calc.loc[fair_share_price_calc["Metric"] == "Last Dividend", "Value1"]).values[0]
future_dividend = pd.to_numeric(fair_share_price_calc.loc[fair_share_price_calc["Metric"] == "Future Divident", "Value1"]).values[0]
dividend_growth_rate = pd.to_numeric(fair_share_price_calc.loc[fair_share_price_calc["Metric"] == "Divident GR", "Value1"]).values[0]

# Calculate the fair share price using DDM
fair_share_price = (future_dividend * (1 + dividend_growth_rate)) / (calculated_coe - dividend_growth_rate)

print(f"Calculated Fair Share Price: ${fair_share_price:.2f}")
