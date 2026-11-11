"""Download SPUT market data (TSX U-UN.TO in CAD, OTC SRUUF in USD) into data/raw/sput/.

The raw U-UN.TO file is kept exactly as Yahoo serves it, including the Uranium
Participation Corp history spliced onto the ticker before 2021-07-19; the split
into UPC and SPUT segments happens in clean_data.py, never in the raw file.

Not retrievable here: SPUT NAV / units / pounds held (Sprott's terms of use
prohibit scraping; history only via SEDAR+ filings, entered manually).
"""
from fetchers import download_group

if __name__ == "__main__":
    print(download_group("sput").to_string(index=False))
