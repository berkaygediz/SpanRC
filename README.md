# SolidSheets - A Lightweight Spreadsheet Editor

![Screenshot](images/solidsheets_banner_0.png)

![GitHub release (latest by date)](https://img.shields.io/github/v/release/berkaygediz/solidsheets)
![GitHub repo size](https://img.shields.io/github/repo-size/berkaygediz/solidsheets)
![GitHub](https://img.shields.io/github/license/berkaygediz/solidsheets)

SolidSheets is a lightweight spreadsheet editor written in Python, utilizing the PySide6 library for its graphical interface. It offers a straightforward alternative to traditional spreadsheet applications, emphasizing simplicity and efficiency.

## Features

- [x] **File Operations**: Open and save CSV, SSFS (SolidSheets Files) & XLSX (partial).
- [x] **Printing**: Print or export tables to PDF format.
- [x] **Editing**: Modify tables with options to delete, edit rows, and columns.
- [x] **Formula Support**: Includes functions like avg, sum, min, max, count, similargraph, etc.
- [x] **Customizable Toolbar**: Tailor the interface to your needs.
- [x] **Performance**: Fast and lightweight with threading support.
- [x] **User Experience**: Alerts for unsaved changes and supports dark/light mode.
- [x] **Real-Time Statistics**: Displays live updates on row count, column count, and cell count.
- [x] **Multilingual**: Available in English, Turkish, German, Spanish, Azerbaijani.
- [x] **Cross-Platform**: Compatible with Windows, macOS, and Linux.
- [x] **Efficiency**: Designed for power-saving and utilizes hardware acceleration.

## Prerequisites

- Python 3.12+
- PySide6
- matplotlib
- psutil
- pandas
- openpyxl

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/berkaygediz/SolidSheets.git
   ```

2. Install requirements:

   ```bash
   pip install -r requirements.txt
   ```

3. Creating a executable file (Unsigned):

   ```bash
   pyinstaller --noconfirm --onefile --windowed --icon "solidsheets_icon.ico" --name "SolidSheets" --clean --optimize "2" --add-data "solidsheets_icon.png;."  "SolidSheets.py"
   ```

## Usage

Launch SolidSheets from the command line:

```bash
python SolidSheets.py
```

## Contributing

Contributions to the SolidSheets project are welcomed. Please refer to [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute and our code of conduct.

## License

This project is licensed under GNU GPLv3, GNU LGPLv3, and Mozilla Public License Version 2.0.
