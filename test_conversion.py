"""
Test script to verify the CSV to Excel conversion functionality
"""

import pandas as pd
from pathlib import Path

def test_conversion():
    # Read the sample CSV
    sample_file = "sample_data.csv"
    
    print("Testing CSV to Excel conversion...")
    print(f"Reading file: {sample_file}")
    
    # Read CSV
    df = pd.read_csv(sample_file)
    print(f"Original data shape: {df.shape}")
    print("Original columns:", df.columns.tolist())
    print("\nOriginal data:")
    print(df.head())
    
    # Add lock id column (hex without 0x prefix)
    df['lock id'] = df['lock'].apply(lambda x: hex(int(x))[2:] if pd.notna(x) else '')
    
    # Remove the original lock column and reorder
    df = df.drop('lock', axis=1)
    columns = ['lock id', 'fsu', 'start', 'end']
    df = df[columns]
    
    print(f"\nProcessed data shape: {df.shape}")
    print("New columns:", df.columns.tolist())
    print("\nProcessed data with hex conversion:")
    print(df.head())
    
    # Save to Excel
    output_file = "test_output.xlsx"
    df.to_excel(output_file, index=False)
    print(f"\nSuccessfully saved to: {output_file}")
    
    return True

if __name__ == "__main__":
    try:
        test_conversion()
        print("\n✅ Test completed successfully!")
    except Exception as e:
        print(f"\n❌ Test failed: {e}")