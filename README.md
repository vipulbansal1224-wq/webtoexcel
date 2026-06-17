# Web to Excel Extractor

Extract any website's content (Menus, Headings, Images, Text) into a structured Excel file with actual embedded image thumbnails.

## Features
- Works on both static and dynamic (React/Angular/JS) websites using Playwright
- Extracts: Menu Links, Logos, Slider/Banner Images, Headings, Paragraphs, General Links
- Embeds actual image thumbnails in Excel Column D
- Web UI built with Flask

## How to Use

### 1. Install dependencies
```bash
python -m venv venv
.\venv\Scripts\pip install -r requirements.txt
.\venv\Scripts\playwright install chromium
```

### 2. Run Web UI
```bash
.\venv\Scripts\python.exe app.py
```
Open: http://127.0.0.1:5000

### 3. Or use Command Line
```bash
.\venv\Scripts\python.exe extractor.py
```

## Excel Output Columns
| Column | Content |
|--------|---------|
| A | Element Type (Menu Link, Logo, Image, Heading, Paragraph) |
| B | Content (Text or Image URL) |
| C | Extra Details (href, alt text) |
| D | Image Thumbnail (embedded) |
