<div align="center">
  <img src="assets/flickersuite-icon.png" alt="FlickerSuite app icon" width="180" />
  <h1>FlickerSuite</h1>
  <p><strong>Capture a thought. Make sense of it.</strong></p>
  <p>A lightweight desktop duo for capturing thoughts and organizing them in Excel.</p>
</div>

**FlickerDrop** makes it quick to jot down an idea. **FlickerView** helps you browse, search, and edit your notes in a familiar Excel workbook.

## Why this exists

Good ideas often show up when you are busy. FlickerSuite makes it easy to capture them right away, then return later to review and organize them. Your notes stay in an Excel workbook, so they are easy to open, edit, and keep with your files.

## Install and run

### Use the ready-to-run apps

1. Download [`FlickerSuite_v1.0.zip`](https://github.com/AhmedElBehery-bs/FlickerFlow/releases/tag/v1.0.0).
2. Extract the ZIP file.
3. Run `FlickerDrop.exe` to capture thoughts or `FlickerView.exe` to browse and organize your workbook.

Python is not required for the packaged apps.

> **Windows SmartScreen:** Windows may warn that the apps are not signed by a major publisher. If you trust the download, select **More info** → **Run anyway**.

### Run from source

Install Python 3.10 or later, then open a terminal in the `FlickerFlow/src` directory and install the dependencies:

```bash
python -m pip install customtkinter openpyxl PySide6
```

Start either app from that directory:

```bash
python FlickerDrop.pyw
python FlickerView.pyw
```

`Sample_Thoughts.xlsx` is in the same directory as the source scripts, as required by the apps.

## Features

### FlickerDrop
- Quickly save thoughts with **Enter**; use **Ctrl+Enter** for a new line.
- Automatically adds an ID and timestamp to each entry.
- Review recent entries and click one to copy it.
- Organize entries across workbook sheets, with undo support and dark/light themes.

### FlickerView
- Browse, edit, search, and sort workbook data.
- Choose from 11 built-in themes.
- Work across multiple sheets, with customizable tab colors.
- Use undo/redo and safe saving to help protect your changes.

## Tech stack

- **FlickerDrop UI:** CustomTkinter
- **FlickerView UI:** PySide6
- **Excel workbook support:** openpyxl
- **Language:** Python 3.10+

## Releases

See the [Releases page](https://github.com/AhmedElBehery-bs/FlickerFlow/releases) for downloads and release notes. The current release is [v1.0.0](https://github.com/AhmedElBehery-bs/FlickerFlow/releases/tag/v1.0.0).
