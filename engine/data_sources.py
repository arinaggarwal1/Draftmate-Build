# engine/data_sources.py

import csv
import io
import re
import ssl
import urllib.parse
import urllib.request
from typing import List, Dict, Tuple

import certifi


GS_HOST = "docs.google.com"


# -------------------------------------------------
# Public API
# -------------------------------------------------

def load_csv(path: str) -> Tuple[List[Dict[str, str]], List[str]]:
    """
    Load and normalize rows from a local CSV file.

    Returns:
        rows: list of normalized row dictionaries
        headers: ordered list of lowercase column headers
    """
    rows = []

    with open(path, mode="r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)

        headers = [str(h).strip().lower() for h in (reader.fieldnames or [])]

        for row in reader:
            norm = {
                str(k).strip().lower(): str(v).strip()
                for k, v in (row.items() if row else [])
            }
            rows.append(norm)

    return rows, headers


def load_google_sheet(sheet_url: str, timeout: int = 20) -> Tuple[List[Dict[str, str]], List[str]]:
    """
    Load and normalize rows from a public Google Sheets URL.

    Accepts standard sharing links with optional gid.

    Returns:
        rows: list of normalized row dictionaries
        headers: ordered list of lowercase column headers
    """
    export_url = _gsheet_to_export_csv_url(sheet_url)
    if not export_url:
        raise ValueError("Invalid Google Sheets URL")

    csv_text = _fetch_gsheet_csv_text(export_url, timeout=timeout)
    return _parse_csv_text(csv_text)


# -------------------------------------------------
# Internal helpers
# -------------------------------------------------

def _parse_csv_text(csv_text: str) -> Tuple[List[Dict[str, str]], List[str]]:
    rows = []

    f = io.StringIO(csv_text)
    reader = csv.DictReader(f)

    headers = [str(h).strip().lower() for h in (reader.fieldnames or [])]

    for row in reader:
        norm = {
            str(k).strip().lower(): str(v).strip()
            for k, v in (row.items() if row else [])
        }
        rows.append(norm)

    return rows, headers


def _gsheet_to_export_csv_url(url: str) -> str:
    try:
        parsed = urllib.parse.urlparse(url)
    except Exception:
        return ""

    if GS_HOST not in parsed.netloc:
        return ""

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9\-_]+)", parsed.path)
    if not m:
        return ""

    ssid = m.group(1)

    gid = "0"
    if parsed.fragment:
        mg = re.search(r"gid=(\d+)", parsed.fragment)
        if mg:
            gid = mg.group(1)

    q = urllib.parse.urlencode({"format": "csv", "gid": gid})
    return f"https://{GS_HOST}/spreadsheets/d/{ssid}/export?{q}"


def _fetch_gsheet_csv_text(export_csv_url: str, timeout: int = 20) -> str:
    req = urllib.request.Request(
        export_csv_url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/122.0 Safari/537.36"
            )
        },
        method="GET",
    )

    context = ssl.create_default_context(cafile=certifi.where())

    with urllib.request.urlopen(req, timeout=timeout, context=context) as resp:
        if resp.status != 200:
            raise RuntimeError(f"HTTP {resp.status}")

        data = resp.read()

        try:
            return data.decode("utf-8-sig")
        except Exception:
            return data.decode("utf-8", errors="replace")
