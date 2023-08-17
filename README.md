
# KMatrixKrawler

A tool to search through K-Matrix Excel files for specific signals.

## Description

KMatrixKrawler is a Python-based GUI application designed to search through K-Matrix Excel files within the K-Matrix directory. Built using the power of `tkinter` for the graphical interface and `openpyxl` for Excel file operations, it offers a simple yet effective solution for your signal search needs.
Use case: You know a signal exists but want to find out the owner or the origin for example.

## Features

- **Directory Search**: Recursively searches all Excel files in the K-Matrix directory for a specified signal.
- **Real-time Feedback**: Provides a status display and progress bar during the search.
- **Detailed Results**: Outputs the file path, sheet name, and cell row for every match found.

## Prerequisites

- Python 3.x installed on your system.
- `pip` for package management.
- Familiarity with the command-line or terminal usage.
- Access to K-Matrix directory: "S:/EE_Elektrik_Elektronik/Vernetzungsdaten/V000_Verbundrelease/E3 1.2_P/Aktuell/K-Matrix"

## Installation

1. Ensure Python is installed on your system.
2. Clone this repository or download the source code.
3. Navigate to the project directory in your terminal or command prompt.
4. Install required packages:

```bash
pip install -r requirements.txt
```

## Usage

1. Launch the application:

```bash
python KMatrixKrawler.py
```

2. In the GUI:
- Choose the directory containing the Excel files.
- Input the search term.
- Click the "Search" button.

3. Review the results displayed in the GUI.

## Troubleshooting
