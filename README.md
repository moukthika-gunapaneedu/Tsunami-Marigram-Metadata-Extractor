# Tsunami-Marigram-Metadata-Extractor

This is a Python-based tool designed to automatically extract, parse, and structure metadata from historical tsunami marigram records. These marigrams are often stored as TIFF images and contain critical tide gauge information such as latitude, longitude, event date, and comments. This project aims to make these records more discoverable, structured, and ready for further scientific analysis.

## Features
-  Extracts text from marigram TIFF images using OCR (Tesseract).
-  Cleans, normalizes, and parses metadata into structured formats.
-  Handles multiple latitude/longitude formats (decimal, signed, DMS).
-  Detects and standardizes event dates from handwritten or printed marigrams.
-  Outputs metadata into CSV/Excel for downstream research.
-  Includes regex-based parsing patterns for robust extraction.
-  Designed for extensibility to accommodate additional metadata fields.

## How It Works
1. Input raw TIFF marigram scans.
2. Run OCR (Tesseract) to extract text from images.
3. Apply regex-based patterns to detect latitude, longitude, event dates, and comments.
4. Normalize values into consistent formats (decimal degrees, ISO 8601 dates).
5. Save results into a structured dataset (CSV/Excel).

## Installation
```bash
git clone https://github.com/<your-username>/Tsunami-Marigram-Metadata-Extractor.git
cd Tsunami-Marigram-Metadata-Extractor
pip install -r requirements.txt


# Usage
python extract_metadata.py --input ./data/marigrams/ --output ./output/metadata.csv
