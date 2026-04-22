# Document Field Batch Replacement Tool

**[🇺🇸 English Version](README_EN.md)** | **[🇨🇳 中文版](README.md)**

**v0.1 | by keloder | Python 3.8.20**

A web-based document batch field replacement tool supporting Word (.docx/.doc) and WPS formats, with forward replacement and reverse restoration capabilities.

## Features

- **Forward Replacement** — Batch replace field content in documents according to rules
- **Reverse Restoration** — Restore replaced content back to original fields
- **Multi-file Processing** — Support uploading multiple .docx / .doc / .wps files simultaneously
- **Rule Management** — Add, edit, delete, sort, and clear replacement rules
- **Import/Export** — Support JSON and TXT format rule import/export
- **Batch Download** — Download single file or download all results as a package after processing
- **Incognito Mode** — Support incognito mode without retaining any data
- **Theme Switching** — Support light/dark theme switching
- **Multi-language** — Support Chinese/English interface switching
- **Portable Mode** — Configuration file saved in software directory for portability

## Project Structure

```
word_test/
├── web_server.py          # Flask Web server main program
├── replacer.py            # Core replacement logic (forward/reverse)
├── document_handler.py    # Document read/write handler
├── build_exe.py           # PyInstaller packaging script
├── build.bat              # One-click packaging batch script (supports onedir/onefile)
├── templates/
│   ── index.html         # Frontend page
├── static/
│   └── favicon.png        # Web icon
├── icon/
│   └── TH.png             # Application icon
└── requirements.txt       # Dependency list
```

## Quick Start

### Development Mode

```bash
python web_server.py
```

Visit http://127.0.0.1:5000

### Package as EXE

Double-click `build.bat`, select packaging mode:

- **1. onedir** — Directory mode, generates multiple files, faster startup
- **2. onefile** — Single file mode, generates single exe, easier distribution

Wait a few minutes, output files will be generated in `dist/` directory.

### Install Dependencies

```bash
pip install flask python-docx lxml pyinstaller pyinstaller-hooks-contrib psutil
```

## Usage

1. Add replacement rules in the left panel (original text → replacement text)
2. Upload document files to process (drag and drop supported)
3. Select "Forward Replacement" or "Reverse Restoration" mode
4. Click start processing and wait for completion
5. Download processed files

## Data Directory

- **Default Data Directory**: Software directory
- **Configuration File Location**: `config.json` in software directory
- **Portable Mode**: All configuration and data saved in software directory for portability
- **Custom Directory**: Can modify data directory path in settings

## Incognito Mode

- Click incognito mode button to enable
- No uploaded files or processing results will be retained in this mode
- Click again to exit incognito mode

## Tech Stack

- Backend: Python 3.8 + Flask
- Frontend: Native HTML/CSS/JavaScript
- Document Processing: python-docx + lxml
- Packaging: PyInstaller 5.13.2

## Dependencies

| Library | Purpose |
|---|---|
| flask | Web framework |
| python-docx | Word document processing |
| lxml | XML parsing |
| pyinstaller | Packaging tool |
| psutil | Process management |
