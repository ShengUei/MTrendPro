# Stock Information Crawler

## Project Description
This is an automated stock information retrieval tool that uses the Yahoo Finance API to obtain price and market capitalization data for stocks, ETFs, and cryptocurrencies, and stores the results in an Excel file.

## Features
- Read stock symbols from an Excel file
- Retrieve current stock prices and market capitalization using Yahoo Finance API
- Intelligently identify asset types (stocks, ETFs, cryptocurrencies, exchange rates)
- Selectively fetch market capitalization based on asset type
- Write retrieved data to Excel file with date-based columns
- Support for external configuration file for easy setup

## Requirements
- Python 3.6+
- pandas
- yfinance
- openpyxl

## Installation
1. Clone this repository or download the source code
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage
1. Ensure you have a properly formatted Excel file with a column named "Symbol" containing all stock symbols
2. Configure the `config.ini` file:
   ```ini
   [Excel]
   file_path = Your Excel File Path.xlsx
   sheet_name = Your Sheet Name
   
   [Output]
   verbose = True
   ```
3. Run the program:
   ```
   python main.py
   ```
4. The program will fetch the closing price and market capitalization (if applicable) for each stock and write the results to the Excel file

## Configuration Options
### [Excel] Section
- `file_path`: Path to the Excel file
- `sheet_name`: Name of the worksheet containing stock symbols

### [Output] Section
- `verbose`: Whether to output detailed logs (True/False)

## Asset Type Handling
- Stocks (Equity/Stock): Fetch price and market capitalization
- Cryptocurrencies (Cryptocurrency): Fetch price and market capitalization
- ETFs/Mutual Funds: Fetch price only, skip market capitalization
- Exchange Rates (Currency): Fetch price only, skip market capitalization

## Notes
- Ensure your Excel file is not open in another program when running the script
- Yahoo Finance API may limit excessive requests, avoid running the program too frequently in a short period
- Some stock symbols may not be compatible with Yahoo Finance, ensure you use the correct symbol format 