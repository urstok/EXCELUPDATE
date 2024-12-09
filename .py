pip install xlsxwriter
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import re
import os

# Function to fetch historic stock data
def fetch_historic_data(stock, start_date, end_date):
    try:
        data = yf.download(stock, start=start_date, end=end_date)
        if not data.empty:
            return data
        else:
            print(f"No data available for {stock} in the given range.")
            return None
    except Exception as e:
        print(f"Error fetching data for {stock}: {e}")
        return None

# Function to calculate support and resistance levels for a given window
def calculate_support_resistance(data, window=20):
    support_level = data['Low'].rolling(window=window).min().iloc[-1]
    resistance_level = data['High'].rolling(window=window).max().iloc[-1]
    return support_level, resistance_level

# Function to calculate support and resistance for multiple trading styles
def calculate_support_resistance_by_style(data):
    styles = {
        'Swing': 50,
        'Intraday': 10,
        'Long Term': 200,
        'Momentum': 30,
        'Scalping': 5
    }

    support_resistance = {}

    for style, window in styles.items():
        if len(data) >= window:
            support, resistance = calculate_support_resistance(data, window)
            support_resistance[f'Support {style}'] = support
            support_resistance[f'Resistance {style}'] = resistance
        else:
            support_resistance[f'Support {style}'] = None
            support_resistance[f'Resistance {style}'] = None

    return support_resistance

# Function to extract numeric values from a string
def extract_numeric_value(text):
    match = re.search(r'[-+]?[0-9]*\.?[0-9]+', str(text))
    if match:
        try:
            return float(match.group(0))
        except ValueError:
            return None
    return None

# Main function to process stock data and clean data
def process_stock_data(input_file, output_file, start_date, end_date):
    stock_list = pd.read_excel(input_file, sheet_name='Sheet1')['STOCK NAME']
    results = {}

    # Fetch stock data and calculate support/resistance
    for stock in stock_list:
        data = fetch_historic_data(stock, start_date, end_date)

        if data is not None:
            stock_ticker = yf.Ticker(stock)
            latest_price = stock_ticker.history(period="1d")['Close'].iloc[-1]
            support_resistance = calculate_support_resistance_by_style(data)
            support_resistance['Current Price'] = latest_price  # Add latest price
            results[stock] = support_resistance

    # Create a DataFrame for each trading style and save in an Excel file
    with pd.ExcelWriter(output_file) as writer:
        for style in ['Swing', 'Intraday', 'Long Term', 'Momentum', 'Scalping']:
            rows = []
            for stock, values in results.items():
                rows.append({
                    'Stock Name': stock,
                    'Current Price': values['Current Price'],
                    f'Support {style}': values.get(f'Support {style}'),
                    f'Resistance {style}': values.get(f'Resistance {style}')
                })
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=style, index=False)

    # Load the Excel file for cleaning
    xls = pd.ExcelFile(output_file)
    sheets_dict = pd.read_excel(output_file, sheet_name=None)

    # Clean the data by extracting numeric values
    for sheet_name, df in sheets_dict.items():
        for col in df.columns:
            if 'Support' in col or 'Resistance' in col:
                df[col] = df[col].apply(extract_numeric_value)

        # Save the cleaned DataFrame back into the dictionary
        sheets_dict[sheet_name] = df

    # Save the cleaned data to a new Excel file
    cleaned_file_path = 'final_' + output_file
    with pd.ExcelWriter(cleaned_file_path, engine='xlsxwriter') as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# Calculate the date one year ago
end_date = datetime.today().strftime('%Y-%m-%d')  # Today's date
start_date = (datetime.today() - timedelta(days=365)).strftime('%Y-%m-%d')  # One year ago

# User input for file paths
input_file = "STOCK EXCEL.xlsx"  # Input Excel file containing stock names
output_file = "support_resistance_data.xlsx"  # Output Excel file

# Run the script
process_stock_data(input_file, output_file, start_date, end_date)

import pandas as pd

# Load the Excel file
file_path = 'final_support_resistance_data.xlsx'

# Load all sheets from the Excel file into a dictionary of DataFrames
sheets_dict = pd.read_excel(file_path, sheet_name=None)

# Function to calculate 5% range from the current price
def calculate_near_support_resistance(current_price, support, resistance):
    support_threshold = current_price * 0.05
    resistance_threshold = current_price * 0.05

    if abs(current_price - support) <= support_threshold:
        near_support = 'Near Support'
    else:
        near_support = 'Neutral'

    if abs(current_price - resistance) <= resistance_threshold:
        near_resistance = 'Near Resistance'
    else:
        near_resistance = 'Neutral'

    return near_support, near_resistance

# Process each sheet in the dictionary
for sheet_name, df in sheets_dict.items():
    if 'Current Price' in df.columns and f'Support {sheet_name}' in df.columns and f'Resistance {sheet_name}' in df.columns:
        df['Near Support'] = df.apply(lambda row: calculate_near_support_resistance(row['Current Price'], row[f'Support {sheet_name}'], row[f'Resistance {sheet_name}'])[0], axis=1)
        df['Near Resistance'] = df.apply(lambda row: calculate_near_support_resistance(row['Current Price'], row[f'Support {sheet_name}'], row[f'Resistance {sheet_name}'])[1], axis=1)

# Save the updated data to a new Excel file
output_file = 'todaySTOCK.xlsx'  # Keep only this file
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    for sheet_name, df in sheets_dict.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Clean up other generated files
os.remove('support_resistance_data.xlsx')
os.remove('final_support_resistance_data.xlsx')

print(f"Updated data saved to: {output_file}")
