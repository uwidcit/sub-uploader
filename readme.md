# Submission Uploader

A tool for uploading myelearning submissions to google driver and linking each file into your gradebook in googlesheets.

## Installing Dependencies

```bash
pip install -r requirements.txt
```

## Building

```bash
pyinstaller --onefile --windowed --add-data "credentials.json;." app.py
```