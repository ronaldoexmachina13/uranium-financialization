"""Download the FIRST-PUBLISHED value of each revised macro series from ALFRED
(the St. Louis Fed archive of past FRED releases) into data/raw/alfred/.

Why: data/raw/fred/ holds today's revised figures. A forecast made at a past
month-end may only use what was public then. This script stores, for every
month (or quarter), the value as it was first printed and the date it was printed.

Run from the repo root, in the VS Code terminal:
    python scripts/fetch_alfred_first_release.py

Needs a free FRED API key, kept in a file called .env in the repo root
(one line:  FRED_API_KEY=your_key_here ). .env is in .gitignore, so the key
is never committed or pushed.

Uses only the Python standard library. Every run writes NEW dated files and
never overwrites an existing raw file (rule D-005). Nothing else in the repo changes.
"""
import datetime as dt
import hashlib
import json
import os
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "alfred"
API = "https://api.stlouisfed.org/fred/series/observations"

# Series whose figures are revised after first publication (decision of 20 Nov 2026).
SERIES = {
    "CPIAUCSL": "US consumer price index, all urban consumers, seasonally adjusted (monthly)",
    "INDPRO": "US industrial production index (monthly)",
    "GDPC1": "US real GDP, level (quarterly)",
    "A191RL1Q225SBEA": "US real GDP growth, annualised (quarterly)",
    "OECDPRINTO01GYSAM": "OECD total industrial production, growth on a year earlier (monthly)",
    "OECDPRINTO01IXOBSAM": "OECD total industrial production, index level (monthly)",
    "G7PRINTO01IXOBSAM": "G7 total industrial production, index level (monthly)",
}


def read_api_key():
    """Read FRED_API_KEY from the environment, or from the .env file in the repo root."""
    key = os.environ.get("FRED_API_KEY")
    env_file = ROOT / ".env"
    if not key and env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("FRED_API_KEY="):
                key = line.split("=", 1)[1].strip()
    if not key:
        raise SystemExit("No FRED_API_KEY found. Put the line FRED_API_KEY=... in a file called .env in the repo root.")
    return key


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def fetch_first_release(series_id, key):
    """Ask ALFRED for the initial release of every observation (output_type=4).
    The field realtime_start is then the date that value was first published."""
    params = {
        "series_id": series_id,
        "api_key": key,
        "file_type": "json",
        "output_type": "4",          # 4 = initial release only
        "realtime_start": "1776-07-04",  # the earliest date ALFRED accepts: all vintages
        "realtime_end": "9999-12-31",
    }
    url = API + "?" + urllib.parse.urlencode(params)
    with urllib.request.urlopen(url, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


def main():
    key = read_api_key()
    OUT.mkdir(parents=True, exist_ok=True)
    stamp = dt.date.today().strftime("%Y%m%d")
    for series_id, label in SERIES.items():
        try:
            data = fetch_first_release(series_id, key)
        except Exception as e:  # report and go on with the next series
            print(f"{series_id}: FAILED ({e})")
            continue
        obs = data.get("observations", [])
        json_path = OUT / f"{series_id}_initial_release_fetched_{stamp}.json"
        csv_path = OUT / f"{series_id}_initial_release_fetched_{stamp}.csv"
        if json_path.exists() or csv_path.exists():
            print(f"{series_id}: files for today already exist, skipped (raw files are never overwritten)")
            continue
        # 1) the raw answer exactly as ALFRED sent it, without the API key
        data.pop("api_key", None)
        json_path.write_text(json.dumps(data, indent=1), encoding="utf-8")
        # 2) a simple table: period, first published value, date first published
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            f.write("date,value_first_published,first_published_on\n")
            for o in obs:
                f.write(f"{o['date']},{o['value']},{o['realtime_start']}\n")
        first = obs[0]["date"] if obs else "none"
        last = obs[-1]["date"] if obs else "none"
        print(f"{series_id}: {len(obs)} observations, {first} to {last}  ({label})")
        print(f"    {csv_path.relative_to(ROOT)}  sha256 {sha256_of(csv_path)}")
        print(f"    {json_path.relative_to(ROOT)}  sha256 {sha256_of(json_path)}")
    print("Done. Nothing else in the repo was changed.")


if __name__ == "__main__":
    main()
