#!/bin/bash
echo "FSU CSV to Excel Converter"
echo "=========================="
echo ""

# Change to the directory where this script is located
cd "$(dirname "$0")"

echo "Installing dependencies..."
pip3 install -r requirements.txt
echo ""
echo "Starting application..."
python3 csv_to_excel_converter.py