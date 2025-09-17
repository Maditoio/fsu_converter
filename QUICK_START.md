# Quick Start Guide

## What This App Does
Converts CSV files with format `lock,fsu,start,end` to Excel files and adds a hexadecimal conversion column for the lock values.

## How to Run

### Option 1: Double-click the batch file
- Simply double-click `run_converter.bat`
- This will install dependencies and start the app

### Option 2: Command line
```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python csv_to_excel_converter.py
```

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
lock  | lock hex | fsu     | start                      | end
55062 | 0xd716   | 3-9-2-8 | 2025-09-16 00:00:00 +0000 | 2025-09-16 04:00:00 +0000
```

## Test File
Use `sample_data.csv` to test the application with sample data.