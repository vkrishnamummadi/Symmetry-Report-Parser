# Access Control Data Parser

This repository contains tools for parsing and analyzing access control data from various file formats, including Excel and text files.

## Features

- Parse Excel files to extract and display data.
- Parse structured text files to extract key-value pairs and hierarchical data.
- Debugging outputs to verify data parsing.

## File Structure

- `main.py`: Contains the main parsing functions for Excel and text files.
- `cards.txt` and `Archives/cards.txt`: Example text files with access control data.
- `Staff_List_04-22_25.xlsx` and `SymmetryReport.xlsx`: Example Excel files for testing.
- `README.md`: Documentation for the repository.

## Usage

1. **Parsing Excel Files**:
   Use the `parse_excel_file` function in `main.py` to read and analyze Excel files.
   ```python
   from main import parse_excel_file
   parse_excel_file("Staff_List_04-22_25.xlsx")
   ```