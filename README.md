
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

- Anaconda installed on your system. (regular python doesn't work for CARIAD)
- Use Anaconda Prompt
- Access to K-Matrix directory: `S:/EE_Elektrik_Elektronik/Vernetzungsdaten/V000_Verbundrelease/E3 1.2_P/Aktuell/K-Matrix`

## Installation

1. Clone this repository or download the source code.
2. Navigate to the project directory in your terminal or command prompt.

## Usage

1. Launch the application:

```bash
python KMatrixKrawler.py
```

2. In the GUI:
- Input the signal name.
- Click the "Suchen" button.

3. Review the results displayed in the GUI.

## Contributions

For contributions, please create a pull request or open an issue to discuss proposed changes.

## License

MIT License. See LICENSE file for more information.
