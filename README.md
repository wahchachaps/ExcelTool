# ExcelTool

Desktop XML-to-Excel converter built with PyQt6 + QML, focused on fast single and batch conversion workflows.

## Current Features
- Convert XML to formatted Excel (`.xlsx`) with type-specific layouts/formulas.
- Supported conversion types:
  - `Den`
  - `Glacier`
  - `Globe`
- Single-file conversion flow.
- Batch conversion flow with per-file live status:
  - `Queued`
  - `Processing`
  - `Done`
  - `Failed`
- Batch file selection options:
  - File dialog (multi-select)
  - Drag-and-drop files
  - Drag-and-drop folders (auto-load all `.xml` recursively)
- File list management before conversion:
  - Remove file from batch list (`x` button)
  - Dynamic file list sizing with scroll when needed
- Batch review screen before save:
  - Edit output file names per file
  - Edit output folder per file
  - Browse folder per file
  - Set one folder for all outputs (`Browse` + `Apply All`)
- Overwrite confirmation dialogs for single and batch saves.
- Remembered user preferences (persistent):
  - Last XML type
  - Last open folder
  - Last single-save folder
  - Last batch-save folder

## Requirements
- Python 3.10+
- Windows

Install dependencies:

```bash
pip install pyqt6 pandas pywin32 xlsxwriter
```

## Run

```bash
python main.py
```

## Usage

### Single File
1. Select one `.xml` file.
2. Pick XML type (`Den`, `Glacier`, `Globe`).
3. Click **Confirm and Convert**.
4. Choose save path for `.xlsx`.

### Batch
1. Select multiple `.xml` files or drop a folder containing `.xml` files.
2. (Optional) Remove unwanted files from the list.
3. Pick XML type.
4. Click **Confirm and Convert**.
5. Wait for processing and review outputs in **Batch Review**.
6. (Optional) Use one folder for all outputs, then click **Save All**.

## Notes
- XML parsing expects the `ArrayFieldDataSet` structure used by the current conversion logic.
- If a file fails during batch conversion, processing continues for remaining files and failed files are marked in status.

## Project Structure
- `main.py`: backend logic (selection, conversion, batch processing, saving, settings persistence)
- `main.qml`: frontend UI (states, drag/drop, file list, batch status/review)

## Author
wahchachaps
