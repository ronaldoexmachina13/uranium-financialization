"""Download commodity prices (gold, silver, copper, WTI) into data/raw/commodities/
and uranium prices (Cameco month-end spot, IMF monthly average) into data/raw/uranium/."""
import pandas as pd

from fetchers import download_group

if __name__ == "__main__":
    print(pd.concat([download_group("market"), download_group("uranium")]).to_string(index=False))
