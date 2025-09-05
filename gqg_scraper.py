import requests
import PyPDF2
from io import BytesIO
import pandas as pd
import re

url = "https://gqg.com/content/2025/08/GQG-Partners-Global-Equity-Fund-Factsheet_AU_Jul25.pdf"

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

def parse_gics_table(gics_text):
    lines = gics_text.splitlines()
    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith("Sector"):
            header_idx = i
            break
    if header_idx is None:
        return pd.DataFrame()
    data_lines = lines[header_idx+1:]
    table = []
    for line in data_lines:
        match = re.match(r"^([A-Za-z\s]+)\s+([\d\.\-]+)\s+([\d\.\-]+)\s+([\d\.\-]+)", line.strip())
        if match:
            sector, fund, index, diff = match.groups()
            table.append({
                "Sector": sector.strip(),
                "Fund": float(fund),
                "Index": float(index),
                "-/+": float(diff)
            })
    return pd.DataFrame(table)

def parse_holdings_table(holdings_text):
    lines = holdings_text.splitlines()
    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith("Company"):
            header_idx = i
            break
    if header_idx is None:
        return pd.DataFrame()
    data_lines = lines[header_idx+1:]
    table = []
    for line in data_lines:
        match = re.match(r"^([A-Za-z0-9\s\.\-&]+)\s+([\d\.\-]+)", line.strip())
        if match:
            company, percent = match.groups()
            table.append({
                "Company": company.strip(),
                "%": float(percent)
            })
    return pd.DataFrame(table)

def parse_countries_table(countries_text):
    lines = countries_text.splitlines()
    header_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith("Country"):
            header_idx = i
            break
    if header_idx is None:
        return pd.DataFrame()
    data_lines = lines[header_idx+1:]
    table = []
    for line in data_lines:
        match = re.match(r"^([A-Za-z\s]+)\s+([\d\.\-]+)\s+([\d\.\-]+)\s+([\d\.\-]+)", line.strip())
        if match:
            country, fund, index, diff = match.groups()
            table.append({
                "Country": country.strip(),
                "Fund": float(fund),
                "Index": float(index),
                "-/+": float(diff)
            })
    return pd.DataFrame(table)

# Parse and save GICS Sectors table
if gics_data:
    gics_df = parse_gics_table(gics_data[0])
    if not gics_df.empty:
        gics_df.to_excel("gqg_factsheet_gics.xlsx", index=False)
        print("GICS Sectors Table:")
        print(gics_df.to_string(index=False))

# Parse and save Top 10 Holdings table
if holdings_data:
    holdings_df = parse_holdings_table(holdings_data[0])
    if not holdings_df.empty:
        holdings_df.to_excel("gqg_factsheet_holdings.xlsx", index=False)
        print("\nTop 10 Holdings Table:")
        print(holdings_df.to_string(index=False))

# Parse and save Top 10 Countries table
if countries_data:
    countries_df = parse_countries_table(countries_data[0])
    if not countries_df.empty:
        countries_df.to_excel("gqg_factsheet_countries.xlsx", index=False)
        print("\nTop 10 Countries Table:")
        print(countries_df.to_string(index=False))