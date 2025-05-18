# Data Cleaning Tool: Remove Zero-Sum Rows from Excel Files

## 📌 Overview

This Python tool cleans Excel files by removing rows where pairs of values in the "Amount in local currency" column sum to zero (with a small tolerance for floating-point precision). It processes all sheets in an Excel workbook while maintaining data integrity.

## ✨ Key Features

- Processes multi-sheet Excel files (.xlsx)
- Identifies and removes zero-sum pairs efficiently using a two-pointer algorithm
- Preserves the original sum of the "Amount in local currency" column
- Generates verification reports showing sums before and after cleaning
- Handles floating-point precision with configurable tolerance

## 📦 Requirements

- Python 3.6+
- pandas
- numpy
- openpyxl

Install dependencies with:
```bash
pip install pandas numpy openpyxl
```

## 🛠️ Usage

1. Place your Excel file in a known directory
2. Modify these variables in the script:
   - `file_path`: Path to your input Excel file
   - `output_file`: Path for the cleaned output file
3. Run the script:
```bash
python clean_data.py
```

## 🔧 How It Works

### Cleaning Algorithm
1. **Sorts** the data by the "Amount in local currency" column
2. Uses a **two-pointer technique** to efficiently find pairs that sum to zero:
   - Starts with pointers at both ends of the sorted data
   - Moves pointers inward based on whether the current sum is negative or positive
3. Removes all identified zero-sum pairs
4. **Verifies** the total sum remains unchanged after cleaning

### File Processing
- Processes each sheet in the input Excel file independently
- Only modifies sheets containing the "Amount in local currency" column
- Preserves all other columns and data structure
- Creates a new Excel file with cleaned data

## 📊 Output Verification

The script provides console output showing:
- Sheets processed
- Sum before and after cleaning for each sheet
- Warnings for any sheets that don't contain the target column
- Errors if any sum mismatches occur (should not happen with proper implementation)

## 📂 Example

**Input Data (Sheet1):**
```
ID | Date       | Amount in local currency | Description
1  | 2023-01-01 | 100.00                  | Deposit
2  | 2023-01-02 | -100.00                 | Withdrawal
3  | 2023-01-03 | 50.00                   | Purchase
4  | 2023-01-04 | -50.00                  | Refund
5  | 2023-01-05 | 75.00                   | Service Fee
```

**Output Data (Sheet1):**
```
ID | Date       | Amount in local currency | Description
5  | 2023-01-05 | 75.00                   | Service Fee
```

## ⚙️ Configuration

Adjust these parameters in the script:
- `tolerance = 1e-10`: Controls floating-point comparison precision
- Column name: Change `'Amount in local currency'` to match your data

## 🚨 Limitations

- Requires the target column to exist in sheets you want to process
- Currently only works with Excel files (.xlsx)
- Maintains original row order except for removed pairs

## 📜 License

MIT License - Free for modification and distribution
