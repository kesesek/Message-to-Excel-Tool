# Message-to-Excel Tool

A simple Windows-friendly Python app that converts a delivery message into a row in an Excel workbook.

## Features

- Paste one structured delivery message into the app window
- Extract fields using regex patterns
- Choose a local `.xlsx` file and optional sheet name
- Append values as a new row and save automatically
- Missing fields remain blank instead of causing errors

## Fields extracted

- Company Name
- Product Info
- Purchase No.
- Invoice No.
- Delivery Date
- Vehicle No.
- Phone
- Name
- ID Card No.
- Tonnage

## Install

1. Install Python 3.10+ (Windows recommended)
2. Create and activate a virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
venv\\Scripts\\activate    # Windows
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

If the current Python installation does not support Tk, the tool will fall back to a simple command-line mode.

## Usage

1. Paste the delivery message into the text area.
2. Choose or create a `.xlsx` file.
3. Optionally enter a sheet name; leave blank to use the first tab.
4. Click `Append to Excel`.

## Packaging to `.exe`

Install `pyinstaller` and build the executable:

```bash
pip install pyinstaller
pyinstaller --onefile main.py
```

The `.exe` will be available under `dist/main.exe`.
