#!/usr/bin/env python3
"""
Monthly OHLCV Data Fetcher for Chinese A-Shares
===============================================
Fetches monthly Open, High, Low, Close, Volume data from yfinance
for Chinese A-share stocks listed in an Excel file.

Usage:
    python fetch_monthly_ohlcv.py -i <input_excel> -o <output_csv>
    
Example:
    python fetch_monthly_ohlcv.py -i "All A shares.xlsx" -o "monthly_ohlcv.csv"

Requirements:
    pip install yfinance pandas openpyxl
"""

import pandas as pd
import yfinance as yf
from datetime import datetime
import os
import sys
import argparse
import time
import warnings
warnings.filterwarnings('ignore')

# Configuration
START_DATE = "2010-01-01"
END_DATE = "2025-12-31"
MAX_RETRIES = 3
RETRY_DELAY = 2  # seconds between retries
RATE_LIMIT_DELAY = 1  # seconds between requests


def convert_ticker_format(ticker):
    """
    Convert Chinese A-share ticker format for yfinance compatibility.
    
    yfinance uses:
    - .SS for Shanghai stocks (instead of .SH)
    - .SZ for Shenzhen stocks (same)
    - .BJ for Beijing stocks (same)
    
    Args:
        ticker: Original ticker string (e.g., "600000.SH")
    
    Returns:
        yfinance-compatible ticker (e.g., "600000.SS")
    """
    ticker = str(ticker).strip()
    if ticker.endswith('.SH'):
        return ticker.replace('.SH', '.SS')
    return ticker


def fetch_with_retry(ticker, start_date, end_date, max_retries=MAX_RETRIES):
    """
    Fetch data with retry logic for rate limiting.
    
    Args:
        ticker: yfinance ticker symbol
        start_date: Start date
        end_date: End date
        max_retries: Maximum retry attempts
    
    Returns:
        DataFrame or None
    """
    for attempt in range(max_retries):
        try:
            data = yf.download(
                ticker, 
                start=start_date, 
                end=end_date, 
                progress=False,
                auto_adjust=False
            )
            if not data.empty:
                return data
            time.sleep(RATE_LIMIT_DELAY)
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(RETRY_DELAY * (attempt + 1))
            else:
                return None
    return None


def fetch_monthly_ohlcv(ticker, start_date, end_date):
    """
    Fetch daily data from yfinance and resample to monthly OHLCV.
    
    Monthly aggregation rules:
    - Open: First trading day's open of the month
    - High: Maximum high of the month
    - Low: Minimum low of the month
    - Close: Last trading day's close of the month
    - Volume: Sum of volume for the month
    
    Args:
        ticker: yfinance-compatible ticker symbol
        start_date: Start date string (YYYY-MM-DD)
        end_date: End date string (YYYY-MM-DD)
    
    Returns:
        tuple: (monthly_data DataFrame or None, error message or None)
    """
    try:
        # Download data with retry
        data = fetch_with_retry(ticker, start_date, end_date)
        
        if data is None or data.empty:
            return None, "No data available"
        
        # Check if we have data covering the required period
        data_start = data.index.min()
        data_end = data.index.max()
        required_start = pd.Timestamp(start_date)
        required_end = pd.Timestamp(end_date)
        
        # Allow 30-day buffer for start/end dates
        if data_start > required_start + pd.Timedelta(days=30):
            return None, f"Data starts at {data_start.strftime('%Y-%m-%d')}, required {start_date}"
        if data_end < required_end - pd.Timedelta(days=30):
            return None, f"Data ends at {data_end.strftime('%Y-%m-%d')}, required {end_date}"
        
        # Handle MultiIndex columns from yfinance (when downloading single ticker)
        if isinstance(data.columns, pd.MultiIndex):
            # Find the correct column names
            def find_col(keyword):
                matches = [c for c in data.columns if keyword in c[0]]
                return matches[0] if matches else None
            
            open_col = find_col('Open')
            high_col = find_col('High')
            low_col = find_col('Low')
            close_col = find_col('Close')
            volume_col = find_col('Volume')
            
            if not all([open_col, high_col, low_col, close_col, volume_col]):
                # Try simple column names as fallback
                if all(c in data.columns.get_level_values(0) for c in ['Open', 'High', 'Low', 'Close', 'Volume']):
                    open_col, high_col, low_col, close_col, volume_col = 'Open', 'High', 'Low', 'Close', 'Volume'
                else:
                    return None, "Could not identify price columns"
        else:
            open_col, high_col, low_col, close_col, volume_col = 'Open', 'High', 'Low', 'Close', 'Volume'
        
        # Resample to monthly (Month End)
        monthly_data = pd.DataFrame()
        monthly_data['open'] = data[open_col].resample('ME').first()
        monthly_data['high'] = data[high_col].resample('ME').max()
        monthly_data['low'] = data[low_col].resample('ME').min()
        monthly_data['close'] = data[close_col].resample('ME').last()
        monthly_data['volume'] = data[volume_col].resample('ME').sum()
        
        # Drop months with all NaN values
        monthly_data = monthly_data.dropna(how='all')
        
        if monthly_data.empty:
            return None, "No valid monthly data after resampling"
        
        # Reset index to make date a column
        monthly_data = monthly_data.reset_index()
        monthly_data = monthly_data.rename(columns={'Date': 'date'})
        
        return monthly_data, None
        
    except Exception as e:
        return None, f"Error: {str(e)}"


def process_tickers(input_file, output_file, max_tickers=None):
    """
    Main function to process tickers from Excel file.
    
    Args:
        input_file: Path to Excel file with tickers in column A
        output_file: Path for output CSV file
        max_tickers: Optional limit on number of tickers to process (for testing)
    
    Returns:
        list: Tuples of (dropped_ticker, reason)
    """
    # Read the Excel file
    print(f"Reading tickers from: {input_file}")
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)
    
    # Get tickers from column A (first column)
    tickers = df.iloc[:, 0].dropna().astype(str).tolist()
    
    if max_tickers:
        tickers = tickers[:max_tickers]
        print(f"Processing first {max_tickers} of {len(df.iloc[:, 0].dropna())} tickers")
    else:
        print(f"Found {len(tickers)} tickers to process")
    
    all_data = []
    dropped_tickers = []
    success_count = 0
    
    for i, ticker in enumerate(tickers, 1):
        original_ticker = ticker
        yf_ticker = convert_ticker_format(ticker)
        
        print(f"[{i}/{len(tickers)}] Processing: {original_ticker} (yf: {yf_ticker})", end=" ")
        
        monthly_data, error = fetch_monthly_ohlcv(yf_ticker, START_DATE, END_DATE)
        
        if error:
            print(f"- DROPPED: {error}")
            dropped_tickers.append((original_ticker, error))
        else:
            # Store with original ticker format
            monthly_data['ticker'] = original_ticker
            all_data.append(monthly_data)
            success_count += 1
            print(f"- OK ({len(monthly_data)} records)")
        
        # Rate limiting delay
        time.sleep(RATE_LIMIT_DELAY)
    
    # Combine all data
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Reorder columns: ticker, date, open, high, low, close, volume
        combined_df = combined_df[['ticker', 'date', 'open', 'high', 'low', 'close', 'volume']]
        
        # Format date column
        combined_df['date'] = pd.to_datetime(combined_df['date']).dt.strftime('%Y-%m-%d')
        
        # Sort by ticker and date
        combined_df = combined_df.sort_values(['ticker', 'date'])
        
        # Save to CSV
        combined_df.to_csv(output_file, index=False)
        
        print(f"\n{'='*60}")
        print("SUCCESS SUMMARY")
        print(f"{'='*60}")
        print(f"Output file: {output_file}")
        print(f"Total records: {len(combined_df):,}")
        print(f"Successful tickers: {success_count}")
        print(f"Date range: {combined_df['date'].min()} to {combined_df['date'].max()}")
    else:
        print("\n✗ No data collected!")
    
    # Print summary of dropped tickers
    print(f"\n{'='*60}")
    print(f"DROPPED TICKERS ({len(dropped_tickers)} total)")
    print(f"{'='*60}")
    for ticker, reason in dropped_tickers:
        print(f"  {ticker}: {reason}")
    
    # Also save dropped tickers to a file
    dropped_file = output_file.replace('.csv', '_dropped.txt')
    with open(dropped_file, 'w') as f:
        f.write(f"Dropped Tickers Report\n")
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Total dropped: {len(dropped_tickers)}\n")
        f.write(f"{'='*60}\n\n")
        for ticker, reason in dropped_tickers:
            f.write(f"{ticker}: {reason}\n")
    print(f"\nDropped tickers saved to: {dropped_file}")
    
    return dropped_tickers


def main():
    parser = argparse.ArgumentParser(
        description='Fetch monthly OHLCV data for Chinese A-shares from yfinance',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python fetch_monthly_ohlcv.py -i tickers.xlsx -o output.csv
  python fetch_monthly_ohlcv.py -i tickers.xlsx -o output.csv --max 10
        """
    )
    
    parser.add_argument('-i', '--input', required=True, 
                        help='Input Excel file path (tickers in column A)')
    parser.add_argument('-o', '--output', required=True, 
                        help='Output CSV file path')
    parser.add_argument('--max', type=int, default=None,
                        help='Maximum number of tickers to process (for testing)')
    
    args = parser.parse_args()
    
    # Validate input file exists
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}")
        sys.exit(1)
    
    # Process tickers
    process_tickers(args.input, args.output, args.max)


if __name__ == "__main__":
    main()
