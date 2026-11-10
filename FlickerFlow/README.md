# FlickerSuite (FlickerDrop & FlickerView)

A lightweight desktop combo for capturing random thoughts and organizing them in Excel. 

Instead of locking your notes inside a proprietary app or a database, this suite uses a standard `.xlsx` file as its backend. You get the speed of a quick-capture app with the flexibility of Excel.

The suite includes two separate applications:

1. **FlickerDrop** (Built with CustomTkinter) - A minimal, always-ready text box to quickly jot down a thought and hit `Enter` to save it to your spreadsheet.
2. **FlickerView** (Built with PySide6) - A fully-featured, dark/light themed table viewer to read, edit, search, and organize the thoughts you've saved.

## Download & Run

If you just want to use the apps, you don't need to install Python.

1. Go to the **Releases** tab on the right side of this page.
2. Download `FlickerSuite_v1.0.zip`.
3. Extract the folder anywhere on your PC.
4. Double-click `FlickerDrop.exe` to start writing, or `FlickerView.exe` to manage your data.

*Note: Because these are compiled Python apps and not signed by a major publisher, Windows SmartScreen might show a warning the first time you run them. Click "More info" -> "Run anyway".*

## Features

**FlickerDrop**
* Fast entry: type your thought and hit Enter.
* Multi-line support using Ctrl+Enter.
* Automatically creates a timestamp and ID for every entry.
* "Recent" panel so you can see what you just saved.
* Click-to-copy on recent entries.

**FlickerView**
* 11 built-in themes (Midnight, Nord, Gruvbox, Light, etc.).
* Edit cells directly in the app; it writes straight back to the Excel file.
* Live search and multi-column sorting.
* Multi-sheet support, with custom tab colors.
* Undo/Redo support.
* Safe saving: writes to a temp file first to prevent file corruption.

## Tech Stack

If you're looking at the code, here's what powers it:
* **GUI:** `customtkinter` (FlickerDrop) and `PySide6` (FlickerView).
* **Data:** `openpyxl` for reading and writing the Excel files.
* **Language:** Python 3.10+

## Running from Source

If you want to run the raw `.pyw` files or build the `.exe` files yourself:

1. Clone the repository.
2. Create a virtual environment and install the dependencies:
   ```bash
   pip install -r requirements.txt