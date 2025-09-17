# Quick Start Guide

## What This App Does
Converts CSV files with format `lock,fsu,start,end` to Excel files and replaces the lock column with lock id in hexadecimal format (without 0x prefix).

## How to Run

### Windows
- **Easy way**: Double-click `run_converter.bat`
- **Manual**: `pip install -r requirements.txt` then `python csv_to_excel_converter.py`

### macOS
- **Easy way**: Run `chmod +x run_converter.sh && ./run_converter.sh`
- **Manual**: `pip3 install -r requirements.txt` then `python3 csv_to_excel_converter.py`

### Linux
- **Easy way**: Run `chmod +x run_converter_linux.sh && ./run_converter_linux.sh`
- **Manual**: `pip3 install -r requirements.txt` then `python3 csv_to_excel_converter.py`

## Using the Application

1. **Select Files**: Click "Select CSV Files" and choose your CSV files
2. **Choose Output**: Browse to select where you want the Excel files saved
3. **Convert**: Click "Convert Files" and wait for completion
4. **Open Results**: The app will ask if you want to open the output folder

## Example
**Input CSV:**
```
lock,fsu,start,end
55062,3-9-2-8,2025-09-16 00:00:00 +0000,2025-09-16 04:00:00 +0000
```

**Output Excel:**
```
lock id | fsu     | start                      | end
d716    | 3-9-2-8 | 2025-09-16 00:00:00 +0000 | 2025-09-16 04:00:00 +0000
```

## Test File
Use `sample_data.csv` to test the application with sample data.