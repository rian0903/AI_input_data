# AI Input Data Parser (OCR to Excel)

This project contains a Python script that uses OCR (RapidOCR) to parse and process scanned form images of library data (specifically from school libraries in Bireuen, Aceh) and automatically formats and writes the extracted information into an Excel spreadsheet.

## Features

- **Automated OCR Processing**: Uses `rapidocr_onnxruntime` to parse images.
- **Spell Correction**: Automated cleaning of common OCR typos specifically tailored to Indonesian/Aceh place names and library terms (e.g., correcting *BIREVEN* to *BIREUEN*, *GAMPONO* to *GAMPONG*, etc.).
- **Smart Data Structuring**: Extracts 24 columns of data, including library details, administration records, collections, and service details.
- **Excel Auto-Save**: Writes and updates structured data in `Usulan_BBB_2026_Bireun_agustus.xlsx`.
- **Automatic Archival**: Moves processed images to a `processed/<timestamp>` folder after completion to keep the working folder clean.

## File Structure

- `input_session.py`: The main OCR parsing and Excel generation script.
- `Usulan_BBB_2026_Bireun_agustus.xlsx`: Target Excel template/document.
- `.gitignore`: Configured to exclude Python cache, temporary files, and IDE configs.

## Getting Started

1. Place image files of the forms (e.g., page 1 and page 2) in the root directory.
2. Install the requirements:
   ```bash
   pip install openpyxl rapidocr_onnxruntime
   ```
3. Run the script:
   ```bash
   python input_session.py
   ```
