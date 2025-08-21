#!/usr/bin/env python3
"""
Tsunami Marigram OCR → Excel
-------------------------------------------------
What this script does
 1) OCR marigram (tide‑gauge) images using Tesseract + OpenCV
 2) Parse key metadata (COUNTRY, STATE, LOCATION, DATE, SCALE)
 3) Geocode LAT/LON online from COUNTRY/STATE/LOCATION using OpenStreetMap Nominatim
 4) Write/append structured rows to an Excel workbook 

Rules
- Region codes: ONLY the official IOC list in this file or your --region-map CSV.
- Country/State/Location: If not in the allow-lists you provide, keep the OCR text and let the geocoder resolve it (no guessing beyond that).
- LAT/LON: Numeric decimals from geocoding. Do NOT scrape degrees/letters from the image.
- RECORDED_DATE: Normalized to YYYY/MM/DD from whatever the image shows.
- SCALE: Reads 1/12, 1:12, SCALE 1/12, etc. Normalized to `1:NN`.

Usage (examples):
  python marigram_ocr_to_excel_geocode.py \
      --images /path/to/marigrams \
      --out-xlsx ./Tsunami_Microfilm_Inventory_Output.xlsx \
      --save-ocr ./ocr_texts \
      --region-map ./ioc_region_codes.csv \
      --country-list ./countries.txt \
      --state-list ./states.txt \
      --location-list ./locations.txt

Minimal:
  python marigram_ocr_to_excel_geocode.py --images ./marigrams --out-xlsx ./out.xlsx

Pip deps:
  opencv-python pillow pytesseract pandas openpyxl numpy geopy
System deps:
  - Tesseract binary (Ubuntu: `sudo apt-get install tesseract-ocr`; macOS: `brew install tesseract`).

Notes:
  - Best‑effort parsing; anything uncertain is left blank for review.
  - The script never invents IOC region codes; they must appear in the text or CSV.
"""

from __future__ import annotations
import argparse
import csv
import re
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import cv2  # type: ignore
import numpy as np  # type: ignore
import pandas as pd  # type: ignore
from PIL import Image
import pytesseract  # type: ignore

# Geocoding
try:
    from geopy.geocoders import Nominatim  # type: ignore
    from geopy.extra.rate_limiter import RateLimiter  # type: ignore
except Exception:
    Nominatim = None
    RateLimiter = None

# ---------------------------
# Columns
# ---------------------------
DEFAULT_COLUMNS = [
    "FILE_NAME", "COUNTRY", "STATE", "LOCATION", "LOCATION_SHORT", "REGION_CODE",
    "START_RECORD", "END_RECORD", "TSEVENT_ID", "TSRUNUP_ID", "RECORDED_DATE",
    "LATITUDE", "LONGITUDE", "IMAGES", "SCALE", "MICROFILM_NAME", "COMMENTS",
]

# ---------------------------
# IOC Region Codes (exact list you approved)
# ---------------------------
IOC_REGION_CODES = {
    "30": "Red Sea and Persian Gulf",
    "40": "Black Sea and Caspian Sea",
    "50": "Mediterranean Sea",
    "60": "Indian Ocean (including W. Australia and W. Indonesia)",
    "70": "Southeast Atlantic Ocean",
    "71": "Southwest Atlantic Ocean",
    "72": "Northwest Atlantic Ocean",
    "73": "Northeast Atlantic Ocean",
    "74": "Caribbean Sea and Bermuda",
    "75": "East Coast of United States and Canada",
    "76": "Gulf of America/Mexico",
    "77": "West Coast of Africa",
    "78": "Central Africa",
    "80": "Hawaii, Johnston Atoll, Midway I",
    "81": "E. Australia, New Zealand, South Pacific Is.",
    "82": "New Caledonia, New Guinea, Solomon Is., Vanuatu",
    "83": "E. Indonesia and Malaysia",
    "84": "China, North and South Korea, Philippines, Taiwan",
    "85": "Japan",
    "86": "Kamchatka and Kuril Islands",
    "87": "Alaska (including Aleutian Islands)",
    "88": "West Coast of North and Central America",
    "89": "West Coast of South America",
}

# ---------------------------
# Regexes
# ---------------------------
DATE_PATTERNS = [
    re.compile(r"(?<!\d)(?P<y>19\d{2}|20\d{2})-(?P<m>0[1-9]|1[0-2])-(?P<d>0[1-9]|[12]\d|3[01])(?!\d)"),
    re.compile(r"(?<!\d)(?P<m>0?[1-9]|1[0-2])[\-/](?P<d>0?[1-9]|[12]\d|3[01])[\-/](?P<y>19\d{2}|20\d{2})(?!\d)"),
    re.compile(r"(?<!\w)(?P<d>0?[1-9]|[12]\d|3[01])\s+(?P<mon>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+(?P<y>19\d{2}|20\d{2})(?!\w)", re.I),
]
MONTH_MAP = { 'JAN':'01','FEB':'02','MAR':'03','APR':'04','MAY':'05','JUN':'06','JUL':'07','AUG':'08','SEP':'09','SEPT':'09','OCT':'10','NOV':'11','DEC':'12' }

SCALE_PATTERNS = [  # 1/12, 1:12, SCALE 1/12 → 1:12
    re.compile(r"(?:SCALE\s*[:=]?\s*)?1\s*[:/]\s*(?P<den>\d{1,4})", re.I),
]

UPPER_TRIPLE_SPLIT = re.compile(r"^([A-Z][A-Z\- .'()&/]+?)\s{2,}([A-Z][A-Z\- .'()&/]+?)\s{2,}([A-Z0-9][A-Z0-9\- .,'()&/]+)$")

# ---------------------------
# Data structures
# ---------------------------
@dataclass
class Row:
    FILE_NAME: str
    COUNTRY: str = ""
    STATE: str = ""
    LOCATION: str = ""
    LOCATION_SHORT: str = ""
    REGION_CODE: str = ""
    START_RECORD: str = ""
    END_RECORD: str = ""
    TSEVENT_ID: str = ""
    TSRUNUP_ID: str = ""
    RECORDED_DATE: str = ""
    LATITUDE: str = ""
    LONGITUDE: str = ""
    IMAGES: str = ""
    SCALE: str = ""
    MICROFILM_NAME: str = ""
    COMMENTS: str = ""

# ---------------------------
# IO helpers
# ---------------------------

def read_list(path: Optional[str]) -> List[str]:
    if not path:
        return []
    vals: List[str] = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if s:
                vals.append(s.upper())
    return vals


def read_region_map(path: Optional[str]) -> Dict[str, str]:
    if not path:
        return IOC_REGION_CODES.copy()
    mapping: Dict[str, str] = {}
    with open(path, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        _ = next(reader, None)
        for row in reader:
            if len(row) < 2:
                continue
            code = row[0].strip()
            desc = row[1].strip()
            if code and desc:
                mapping[code] = desc
    return mapping

# ---------------------------
# OCR pipeline
# ---------------------------

def load_image(path: str) -> np.ndarray:
    img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        raise RuntimeError(f"Failed to read image: {path}")
    return img


def preprocess_variants(img: np.ndarray) -> List[np.ndarray]:
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    out: List[np.ndarray] = []
    _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU); out.append(th)
    _, th_inv = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU); out.append(th_inv)
    ad = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, 11); out.append(ad)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8)).apply(gray)
    _, th2 = cv2.threshold(clahe, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU); out.append(th2)
    blur = cv2.GaussianBlur(gray, (3,3), 0)
    _, th3 = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU); out.append(th3)
    return out


def ocr_image(img: np.ndarray) -> Tuple[str, float]:
    config = "--psm 6 --oem 3"
    pil_img = Image.fromarray(img)
    data = pytesseract.image_to_data(pil_img, config=config, output_type=pytesseract.Output.DATAFRAME)
    text = "\n".join([str(t) for t in data["text"].fillna("") if str(t).strip()])
    confs = [c for c in data.get("conf", []).tolist() if isinstance(c, (int, float)) and c >= 0]
    avg_conf = float(np.mean(confs)) if confs else 0.0
    return text, avg_conf


def best_ocr_from_variants(img: np.ndarray) -> Tuple[str, float, np.ndarray]:
    best_text = ""; best_conf = -1.0; best_variant = img
    for var in preprocess_variants(img):
        text, conf = ocr_image(var)
        if conf > best_conf or (conf == best_conf and len(text) > len(best_text)):
            best_text, best_conf, best_variant = text, conf, var
    return best_text, best_conf, best_variant

# ---------------------------
# Parsing helpers
# ---------------------------

def sanitize_text(text: str) -> str:
    text = text.replace("\x0c", " ")
    text = re.sub(r"[\u200b\u200c\u200d]", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text


def parse_country_state_location(lines: List[str], countries: List[str], states: List[str]) -> Tuple[str, str, str]:
    # A) UPPERCASE triple with 2+ spaces
    for line in lines[:15]:
        m = UPPER_TRIPLE_SPLIT.match(line.strip())
        if m:
            return m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
    # B) Semicolon/comma triplets
    for line in lines[:20]:
        parts = re.split(r"\s*[;,\t]\s*", line.strip())
        if len(parts) >= 3:
            a, b, c = parts[0], parts[1], parts[2]
            if a.upper() in countries or a.isupper():
                return a.strip(), b.strip(), c.strip()
    # C) Explicit labels
    blob = "\n".join(lines[:50])
    m = re.search(r"COUNTRY[:\-\s]+([A-Z .,'()&/-]+)", blob, re.I)
    country = m.group(1).strip() if m else ""
    m = re.search(r"STATE[:\-\s]+([A-Z0-9 .,'()&/-]+)", blob, re.I)
    state = m.group(1).strip() if m else ""
    m = re.search(r"LOCATION[:\-\s]+([A-Z0-9 .,'()&/-]+)", blob, re.I)
    location = m.group(1).strip() if m else ""
    return country, state, location


def normalize_date_to_ymd(text: str) -> str:
    for pat in DATE_PATTERNS:
        m = pat.search(text)
        if not m:
            continue
        gd = {k: (v if v is None else str(v)) for k, v in m.groupdict().items()}
        if 'mon' in gd and gd['mon']:
            y = gd['y']; d = gd['d'].zfill(2)
            mon = gd['mon'].upper()[:4].replace('.', '')
            mm = MONTH_MAP.get(mon[:3], '')
            if y and mm and d:
                return f"{y}/{mm}/{d}"
        else:
            y = gd.get('y'); mm = gd.get('m'); d = gd.get('d')
            if y and mm and d:
                return f"{y}/{mm.zfill(2)}/{d.zfill(2)}"
    return ""


def parse_scale(text: str) -> str:
    for pat in SCALE_PATTERNS:
        m = pat.search(text)
        if m:
            return f"1:{m.group('den')}"
    return ""

# ---------------------------
# Geocoding
# ---------------------------

def make_geocoder() -> Optional[RateLimiter]:
    if Nominatim is None:
        return None
    geolocator = Nominatim(user_agent="marigram_geocoder")
    return RateLimiter(geolocator.geocode, min_delay_seconds=1.0)


def geocode_latlon(country: str, state: str, location: str, geocode_fn: Optional[RateLimiter]) -> Tuple[str, str]:
    if geocode_fn is None:
        return "", ""
    queries: List[str] = []
    if location and state and country:
        queries.append(f"{location}, {state}, {country}")
    if location and country:
        queries.append(f"{location}, {country}")
    if state and country:
        queries.append(f"{state}, {country}")
    if country:
        queries.append(country)
    for q in queries:
        try:
            loc = geocode_fn(q)
            if loc and getattr(loc, 'latitude', None) is not None and getattr(loc, 'longitude', None) is not None:
                return f"{float(loc.latitude):.5f}", f"{float(loc.longitude):.5f}"
        except Exception:
            continue
    return "", ""

# ---------------------------
# Main processing
# ---------------------------

def process_image(path: str,
                  region_map: Dict[str, str],
                  countries: List[str],
                  states: List[str],
                  locations: List[str],
                  save_ocr_dir: Optional[Path],
                  geocode_fn: Optional[RateLimiter]) -> Row:
    img = load_image(path)
    text, conf, best_variant = best_ocr_from_variants(img)

    if save_ocr_dir:
        save_ocr_dir.mkdir(parents=True, exist_ok=True)
        (save_ocr_dir / (Path(path).stem + ".txt")).write_text(text, encoding="utf-8")
        cv2.imwrite(str(save_ocr_dir / (Path(path).stem + "_bin.png")), best_variant)

    text_clean = sanitize_text(text)
    lines = [ln.strip() for ln in text_clean.splitlines() if ln.strip()]

    country, state, location = parse_country_state_location(lines, countries, states)

    # Keep OCR values even if not in lists; rely on geocoder to resolve
    # (if not in list, refer to internet to find it.)

    date = normalize_date_to_ymd(text_clean)
    scale = parse_scale(text_clean)

    # Geocode for numeric lat/lon
    lat, lon = geocode_latlon(country, state, location, geocode_fn)

    # REGION_CODE: strict — only accept if explicit code appears and is in mapping
    loc_short = "UNKNOWN"
    region_code = "UNKNOWN"
    m = re.search(r"\b(?:REGION|LOCATION_SHORT|LOC\.? SHORT)[:\s\-\[]+(?P<code>\d{2})\b", text_clean, re.I)
    if m and m.group("code") in region_map:
        region_code = m.group("code")
        loc_short = region_map[region_code]
    else:
        top = " \n".join(lines[:10])
        m2 = re.search(r"\[(?P<code>\d{2})\]", top)
        if m2 and m2.group("code") in region_map:
            region_code = m2.group("code")
            loc_short = region_map[region_code]

    row = Row(
        FILE_NAME=Path(path).name,
        COUNTRY=country,
        STATE=state,
        LOCATION=location,
        LOCATION_SHORT=loc_short,
        REGION_CODE=region_code,
        RECORDED_DATE=date,
        LATITUDE=lat,
        LONGITUDE=lon,
        SCALE=scale,
        IMAGES="1",
        COMMENTS=f"avg_conf={conf:.1f}"
    )
    return row

# ---------------------------
# Excel helpers
# ---------------------------

def gather_images(root: str) -> List[str]:
    exts = {".tif", ".tiff", ".png", ".jpg", ".jpeg", ".webp"}
    paths: List[str] = []
    for p in sorted(Path(root).rglob("*")):
        if p.suffix.lower() in exts:
            paths.append(str(p))
    return paths


def ensure_excel(path: str) -> None:
    p = Path(path)
    if not p.exists():
        df = pd.DataFrame(columns=DEFAULT_COLUMNS)
        df.to_excel(path, index=False)


def append_rows_to_excel(path: str, rows: List[Row]) -> None:
    ensure_excel(path)
    existing = pd.read_excel(path)
    for col in DEFAULT_COLUMNS:
        if col not in existing.columns:
            existing[col] = ""
    new_df = pd.DataFrame([asdict(r) for r in rows], columns=DEFAULT_COLUMNS)
    out = pd.concat([existing, new_df], ignore_index=True)
    out.to_excel(path, index=False)

# ---------------------------
# CLI
# ---------------------------

def main():
    ap = argparse.ArgumentParser(description="OCR marigram images and write Excel rows.")
    ap.add_argument("--images", required=True, help="Folder containing marigram images (tif/png/jpg)")
    ap.add_argument("--out-xlsx", required=True, help="Output Excel path (.xlsx)")
    ap.add_argument("--region-map", default=None, help="CSV with IOC region codes (id,description). If omitted, uses strict list in script.")
    ap.add_argument("--country-list", default=None, help="TXT with known countries (one per line)")
    ap.add_argument("--state-list", default=None, help="TXT with known states/regions (one per line)")
    ap.add_argument("--location-list", default=None, help="TXT with known locations (one per line)")
    ap.add_argument("--save-ocr", default=None, help="Optional folder to save OCR text and binarized image")
    args = ap.parse_args()

    region_map = read_region_map(args.region_map)
    countries = read_list(args.country_list)
    states = read_list(args.state_list)
    locations = read_list(args.location_list)
    save_ocr_dir = Path(args.save_ocr) if args.save_ocr else None

    geocode_fn = make_geocoder()

    paths = gather_images(args.images)
    if not paths:
        print(f"No images found under: {args.images}")
        sys.exit(1)

    rows: List[Row] = []
    for i, path in enumerate(paths, 1):
        try:
            row = process_image(path, region_map, countries, states, locations, save_ocr_dir, geocode_fn)
            rows.append(row)
            print(f"[{i}/{len(paths)}] OK -> {Path(path).name}")
        except Exception as e:
            print(f"[{i}/{len(paths)}] ERROR -> {Path(path).name}: {e}")

    append_rows_to_excel(args.out_xlsx, rows)
    print(f"\nWrote {len(rows)} rows to {args.out_xlsx}")


if __name__ == "__main__":
    main()
