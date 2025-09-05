import requests
import PyPDF2
from io import BytesIO
import pandas as pd
import re
import numpy as np
import sys
import json
from pathlib import Path

with open("global_equity_fund.json", "r") as f:
    fund_data = json.load(f)

# Initialize lists to collect all data
all_gics_data = []
all_holdings_data = []
all_countries_data = []

# Excel files for ranked format only
gics_ranked_excel = "gqg_factsheet_gics_ranked_format.xlsx"
holdings_ranked_excel = "gqg_factsheet_holdings_ranked_format.xlsx"
countries_ranked_excel = "gqg_factsheet_countries_ranked_format.xlsx"

# Track processing order for proper month sequencing
month_order = []

for month, url in fund_data.items():
    if not url or url.strip() in ("", "?"):
        continue
    print(f"Processing {month}: {url}")
    month_order.append(month)

    # Download PDF and read content
    response = requests.get(url)
    pdf_file = BytesIO(response.content)
    reader = PyPDF2.PdfReader(pdf_file)

    gics_data = []
    holdings_data = []
    countries_data = []

    for page in reader.pages:
        text = page.extract_text()
        if "GICS Sectors %" in text:
            start = text.find("GICS Sectors %")
            end = text.find("GQG Partners Global Equity Fund", start)
            gics_text = text[start:end].strip()
            gics_data.append(gics_text)
        if "Top 10 Holdings %" in text:
            start = text.find("Top 10 Holdings %")
            end = text.find("Top 10 Countries %", start)
            holdings_text = text[start:end].strip()
            holdings_data.append(holdings_text)
        if "Top 10 Countries %" in text:
            start = text.find("Top 10 Countries %")
            end = text.find("GICS Sectors %", start)
            countries_text = text[start:end].strip()
            countries_data.append(countries_text)

    def safe_float(val):
        val = val.strip()
        if val in ('-', ''):
            return 0.0  # Treat NaN as 0
        try:
            return float(val)
        except ValueError:
            return 0.0  # Treat NaN as 0

    def parse_gics_table(gics_text, month):
        lines = gics_text.splitlines()
        header_idx = None
        for i, line in enumerate(lines):
            if line.strip().startswith("Sector"):
                header_idx = i
                break
        if header_idx is None:
            return []
        data_lines = lines[header_idx+1:]
        records = []
        
        for line in data_lines:
            line = line.strip()
            print(line)
            match = re.match(
                r"^([A-Za-z0-9\s&/,'\-\.]+)\s+([\d\.\-]+|[-])\s+([\d\.\-]+|[-])\s+([\d\.\-]+|[-])$", 
                line
            )
            if match:
                sector, fund, index, diff = match.groups()
                records.append({
                    "Month": month,
                    "Entity": sector.strip(),
                    "Fund": safe_float(fund),
                    "Index": safe_float(index),
                    "Difference": safe_float(diff)
                })
            else:
                match2 = re.match(
                    r"^([A-Za-z0-9\s&/,'\-\.]+)\s+([\d\.\-]+|[-])\s+([\d\.\-]+|[-])$", 
                    line
                )
                if match2:
                    sector, fund, diff = match2.groups()
                    records.append({
                        "Month": month,
                        "Entity": sector.strip(),
                        "Fund": safe_float(fund),
                        "Index": 0.0,  # Treat missing index as 0
                        "Difference": safe_float(diff)
                    })
        return records

    def parse_holdings_table(holdings_text, month):
        lines = holdings_text.splitlines()
        header_idx = None
        for i, line in enumerate(lines):
            if line.strip().startswith("Holding"):
                header_idx = i
                break
        if header_idx is None:
            return []
        data_lines = lines[header_idx+1:]
        records = []
        for line in data_lines:
            # Fix: allow for company names with commas, apostrophes, and multiple spaces
            match = re.match(r"^([A-Za-z0-9\s\.\-&',/]+)\s+([\d\.\-]+)$", line.strip())
            if match:
                company, percent = match.groups()
                records.append({
                    "Month": month,
                    "Entity": company.strip(),
                    "Holdings_Percent": safe_float(percent)
                })
        return records

    def parse_countries_table(countries_text, month):
        lines = countries_text.splitlines()
        header_idx = None
        for i, line in enumerate(lines):
            if line.strip().startswith("Country"):
                header_idx = i
                break
        if header_idx is None:
            return []
        data_lines = lines[header_idx+1:]
        records = []
        for line in data_lines:
            match = re.match(r"^([A-Za-z\s]+)\s+([\d\.\-]+)\s+([\d\.\-]+)\s+([\d\.\-]+)", line.strip())
            if match:
                country, fund, index, diff = match.groups()
                records.append({
                    "Month": month,
                    "Entity": country.strip(),
                    "Fund": safe_float(fund),
                    "Index": safe_float(index),
                    "Difference": safe_float(diff)
                })
        return records

    # Parse and collect data
    if gics_data:
        gics_records = parse_gics_table(gics_data[0], month)
        all_gics_data.extend(gics_records)
        print(f"GICS Sectors collected: {len(gics_records)} records")

    if holdings_data:
        holdings_records = parse_holdings_table(holdings_data[0], month)
        all_holdings_data.extend(holdings_records)
        print(f"Holdings collected: {len(holdings_records)} records")

    if countries_data:
        countries_records = parse_countries_table(countries_data[0], month)
        all_countries_data.extend(countries_records)
        print(f"Countries collected: {len(countries_records)} records")

def create_ranked_format_df(data_records, value_column, format_as_percent=True):
    """Create a ranked format showing top 10 entries per month with individual month columns"""
    if not data_records:
        return pd.DataFrame()
    
    df = pd.DataFrame(data_records)
    months = [m for m in month_order if m in df['Month'].unique()]
    
    # Create individual DataFrames for each month, then combine horizontally
    month_dfs = []
    
    for month in months:
        month_data = df[df['Month'] == month].copy()
        month_data = month_data.sort_values(value_column, ascending=False).head(10)
        
        # Create month DataFrame with rank
        month_df = pd.DataFrame({
            'Rank': range(1, 11),
            'Month': [month] * 10,
            'Company': [''] * 10,
            'Holding %': [''] * 10
        })
        
        # Fill in the actual data
        for i in range(min(len(month_data), 10)):
            month_df.loc[i, 'Company'] = month_data.iloc[i]['Entity']
            value = month_data.iloc[i][value_column]
            if format_as_percent:
                month_df.loc[i, 'Holding %'] = f"{value:.1f}%" if pd.notna(value) else "0.0%"
            else:
                month_df.loc[i, 'Holding %'] = f"{value:.1f}" if pd.notna(value) else "0.0"
        
        month_dfs.append(month_df)
    
    # Combine all month DataFrames horizontally
    if month_dfs:
        # Start with the first month's Rank column
        result_df = month_dfs[0][['Rank']].copy()
        
        # Add each month's columns with the specified headers
        for i, month_df in enumerate(month_dfs):
            month_cols = ['Month', 'Company', 'Holding %']
            for col in month_cols:
                result_df[f'{months[i]}_{col}'] = month_df[col]
        
        return result_df
    else:
        return pd.DataFrame()

# Create ranked format DataFrames for all data types
print("\nCreating ranked format DataFrames...")

# GICS Sectors ranked format (sorted by Fund percentage, largest to smallest)
if all_gics_data:
    gics_ranked_df = create_ranked_format_df(all_gics_data, 'Fund', format_as_percent=True)
    gics_ranked_df.to_excel(gics_ranked_excel, index=False)
    

# Holdings ranked format (sorted by Holdings_Percent, largest to smallest)
if all_holdings_data:
    holdings_ranked_df = create_ranked_format_df(all_holdings_data, 'Holdings_Percent', format_as_percent=True)
    holdings_ranked_df.to_excel(holdings_ranked_excel, index=False)
    

# Countries ranked format (sorted by Fund percentage, largest to smallest)
if all_countries_data:
    countries_ranked_df = create_ranked_format_df(all_countries_data, 'Fund', format_as_percent=True)
    countries_ranked_df.to_excel(countries_ranked_excel, index=False)
    
# Summary
print(f"\nProcessed {len(month_order)} months: {month_order}")
print(f"Total entities - GICS: {len(set(r['Entity'] for r in all_gics_data))}")
print(f"Total entities - Holdings: {len(set(r['Entity'] for r in all_holdings_data))}")
print(f"Total entities - Countries: {len(set(r['Entity'] for r in all_countries_data))}")
print("\nAll data is now in ranked format showing Top 10 for each month, sorted from largest to smallest.")