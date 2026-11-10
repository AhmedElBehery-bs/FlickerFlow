# FlickerSuite (FlickerDrop & FlickerView)

A lightweight desktop duo for capturing thoughts and keeping them organized.

**FlickerDrop** lets you jot down an idea in seconds. **FlickerView** helps you browse, search, and edit your notes—all saved in a familiar Excel workbook that you can open anytime.

## Download and Run

No Python installation is needed to use the packaged apps:

1. Download [`FlickerSuite_v1.0.zip`](https://github.com/AhmedElBehery-bs/FlickerFlow/releases/tag/v1.0.0) from the latest release.
2. Extract the ZIP file.
3. Open `FlickerDrop.exe` to capture a thought or `FlickerView.exe` to browse and organize your workbook.

> **Windows SmartScreen:** Windows may show a warning because the apps are not signed by a major publisher. If you trust the download, select **More info** → **Run anyway**.

## Features

### FlickerDrop
- Quickly save thoughts with **Enter**; use **Ctrl+Enter** for a new line.
- Automatically adds an ID and timestamp to each entry.
- Review recent entries and click one to copy it.
- Organize entries across workbook sheets, with undo support and dark/light themes.

### FlickerView
- Browse, edit, search, sort, and organize workbook data.
- Choose from 11 built-in themes.
- Work across multiple sheets, with customizable tab colors.
- Use undo/redo and safe saving to help protect your changes.

## Run from Source

The source scripts are in the `FlickerFlow` folder. Install Python 3.10 or later, then install the app dependencies:

```bash
python -m pip install customtkinter openpyxl PySide6
```

Run either app from that folder:

```bash
python FlickerDrop.pyw
python FlickerView.pyw
```

## Tech Stack

- **FlickerDrop UI:** CustomTkinter
- **FlickerView UI:** PySide6
- **Excel workbook support:** openpyxl
- **Language:** Python 3.10+

## Releases

See the [Releases page](https://github.com/AhmedElBehery-bs/FlickerFlow/releases) for downloads and release notes. The current release is [v1.0.0](https://github.com/AhmedElBehery-bs/FlickerFlow/releases/tag/v1.0.0).
