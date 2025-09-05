import requests
import PyPDF2
from io import BytesIO
import pandas as pd

url = "https://gqg.com/content/2025/08/GQG-Partners-Global-Equity-Fund-Factsheet_AU_Jul25.pdf"

# Download PDF and read content
response = requests.get(url)
pdf_file = BytesIO(response.content)
reader = PyPDF2.PdfReader(pdf_file)

# Extract and filter for GICS Sectors section
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

df = pd.DataFrame({
    "GICS Sectors": gics_data,
    "Top 10 Holdings": holdings_data,
    "Top 10 Countries": countries_data
})

print(df)
