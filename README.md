# FSU CSV to Excel Converter

A simple Python GUI application that converts CSV files to Excel format and adds hexadecimal conversion for lock values.

## Features

- **User-friendly GUI** built with tkinter
- **Multi-file processing** - select and convert multiple CSV files at once
- **Lock ID conversion** - replaces lock column with lock id in hexadecimal format (without 0x prefix)
- **Error handling** - validates CSV format and provides detailed error messages
- **Progress tracking** - visual progress bar and status updates
- **Customizable output** - choose where to save converted files
- **Cross-platform** - works on Windows, macOS, and Linux

## Requirements

- Python 3.7 or higher
- pandas
- openpyxl

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Running the Application

**Windows:**
- Double-click `run_converter.bat`, or
- Run: `python csv_to_excel_converter.py`

**macOS:**
- Run: `chmod +x run_converter.sh && ./run_converter.sh`, or
- Run: `python3 csv_to_excel_converter.py`

**Linux:**
- Run: `chmod +x run_converter_linux.sh && ./run_converter_linux.sh`, or
- Run: `python3 csv_to_excel_converter.py`

### Using the GUI

1. **Select CSV Files**: Click "Select CSV Files" to choose one or multiple CSV files
2. **Choose Output Directory**: Browse to select where you want the converted Excel files saved
3. **Convert**: Click "Convert Files" to start the conversion process

## CSV File Format

Your CSV files should have the following format:

```csv
lock,fsu,start,end
55062,3-9-2-8,2025-09-16 00:00:00 +0000,2025-09-16 04:00:00 +0000
55063,3-9-2-9,2025-09-16 04:00:00 +0000,2025-09-16 08:00:00 +0000
```

### Required Columns:
- **lock**: Numeric lock value (will be converted to hex)
- **fsu**: FSU identifier
- **start**: Start timestamp
- **end**: End timestamp

## Output

The application will:
1. Read your CSV files
2. Replace the "lock" column with "lock id" column containing hexadecimal values
3. Convert lock values to hexadecimal without 0x prefix (e.g., 55062 → d716)
4. Save as Excel (.xlsx) files with "_converted" suffix

### Example Output:
```
lock id | fsu     | start                      | end
d716    | 3-9-2-8 | 2025-09-16 00:00:00 +0000 | 2025-09-16 04:00:00 +0000
d717    | 3-9-2-9 | 2025-09-16 04:00:00 +0000 | 2025-09-16 08:00:00 +0000
```

## Error Handling

The application handles various error scenarios:
- **Invalid CSV format**: Missing required columns
- **Invalid lock values**: Non-numeric lock values
- **File access issues**: Permission errors or corrupted files
- **Output directory issues**: Invalid paths or permission problems

## Screenshots

### Main Interface
The application features a clean, intuitive interface with:
- File selection area with list view
- Output directory browser
- Convert button with progress indication
- Status messages and error reporting

## Troubleshooting

### Common Issues:

1. **"Missing required columns" error**
   - Ensure your CSV has exactly these column names: lock, fsu, start, end
   - Check for typos or extra spaces in column headers

2. **"Permission denied" error**
   - Make sure the output directory is writable
   - Close any Excel files that might be open from previous conversions

3. **"Invalid lock value" error**
   - Ensure all lock values are numeric
   - Check for empty cells in the lock column

## Technical Details

- Built with Python's tkinter for cross-platform compatibility
- Uses pandas for efficient CSV/Excel processing
- Implements threading to prevent UI freezing during conversion
- Automatic hexadecimal conversion using Python's built-in hex() function

## License

This project is open source and available under the MIT License.

## Support

If you encounter any issues or have suggestions for improvements, please create an issue in the project repository.