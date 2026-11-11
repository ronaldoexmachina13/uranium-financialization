"""Shared helpers: paths, config, HTTP session, and never-overwrite raw storage.

Raw-file rules
--------------
* Raw files are stored exactly as received (bytes) under data/raw/<raw_dir>/.
* File name: <name>__<source>-<id>__<YYYYMMDD>.<ext> (UTC retrieval date), e.g.
  gold_futures__yahoo-GC-F__20260929.csv. Files renamed to this scheme on
  2026-09-29 are listed (old -> new, SHA-256) in logs/raw_file_renames.csv.
* An existing raw file is NEVER overwritten. If a same-day file exists with the
  same SHA-256 the new download is discarded as "unchanged"; if the content
  differs, the new file gets a _<HHMMSS> suffix.
* Every attempt is appended to logs/retrieval_log.csv.
"""
from __future__ import annotations

import csv
import hashlib
import re
from datetime import datetime, timezone
from pathlib import Path

import requests
import yaml

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
PROCESSED = ROOT / "data" / "processed"
LOGS = ROOT / "logs"
RETRIEVAL_LOG = LOGS / "retrieval_log.csv"
LOG_FIELDS = ["retrieved_at_utc", "key", "source", "identifier", "raw_file",
              "sha256", "bytes", "status", "message"]


def load_config() -> dict:
    with open(ROOT / "config" / "series_config.yaml", encoding="utf-8") as f:
        return yaml.safe_load(f)


def http_session(cfg: dict) -> requests.Session:
    s = requests.Session()
    s.headers["User-Agent"] = cfg["http"]["user_agent"]
    return s


def utc_now() -> datetime:
    return datetime.now(timezone.utc)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def safe_name(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s).strip("-")


def save_raw(data: bytes, raw_dir: str, stem: str, ext: str,
             when: datetime | None = None, dated: bool = True) -> tuple[Path, str]:
    """Store bytes without overwriting. Returns (path, status) where status is
    'saved', 'saved_new_version' or 'unchanged'. With dated=False the stem is
    used as-is (caller supplies e.g. a trade date)."""
    when = when or utc_now()
    folder = RAW / raw_dir
    folder.mkdir(parents=True, exist_ok=True)
    digest = sha256(data)
    base = folder / (f"{safe_name(stem)}__{when:%Y%m%d}.{ext}" if dated else f"{safe_name(stem)}.{ext}")
    same_day = sorted(folder.glob(f"{base.stem}*.{ext}"))
    for p in same_day:
        if sha256(p.read_bytes()) == digest:
            return p, "unchanged"
    path = base if not base.exists() else folder / f"{base.stem}_{when:%H%M%S}.{ext}"
    if path.exists():  # extremely unlikely; refuse rather than overwrite
        raise FileExistsError(path)
    path.write_bytes(data)
    return path, "saved" if path == base else "saved_new_version"


def log_retrieval(key: str, source: str, identifier: str, path: Path | None,
                  data: bytes | None, status: str, message: str = "") -> None:
    LOGS.mkdir(exist_ok=True)
    new = not RETRIEVAL_LOG.exists()
    with open(RETRIEVAL_LOG, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if new:
            w.writeheader()
        w.writerow({
            "retrieved_at_utc": utc_now().isoformat(timespec="seconds"),
            "key": key, "source": source, "identifier": identifier,
            "raw_file": str(path.relative_to(ROOT)) if path else "",
            "sha256": sha256(data) if data is not None else "",
            "bytes": len(data) if data is not None else "",
            "status": status, "message": message,
        })
