# ExcelTool

Desktop XML-to-Excel converter built with PyQt6 and QML.

## Features
- Convert XML files to formatted Excel (`.xlsx`)
- Single-file and batch conversion (select one or multiple XML files)
- Supported conversion types: `Den`, `Glacier`, `Globe`
- Progress/status UI with selected file info (name and size)
- Drag-and-drop support for single XML file

## Requirements
- Python 3.10+ (recommended)
- Windows (uses `pywin32`)

Install dependencies:

```bash
pip install pyqt6 pandas pywin32 xlsxwriter
```

## Run

```bash
python main.py
```

## Usage
1. Click the file selection area and choose one or more `.xml` files.
2. Select conversion type (`Den`, `Glacier`, or `Globe`).
3. Click **Confirm and Convert**.
4. Choose the output `.xlsx` save path.

## Project Structure
- `main.py` - backend logic (selection, processing, save flow)
- `main.qml` - frontend UI (states, controls, layout)

## Author
wahchachaps
