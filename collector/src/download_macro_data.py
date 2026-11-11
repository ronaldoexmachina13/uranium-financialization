"""Phase 3: download macro-financial series (FRED, Cboe VIX, DXY) into data/raw/macro/."""
from fetchers import download_group

if __name__ == "__main__":
    print(download_group("macro").to_string(index=False))
