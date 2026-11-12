You can build a console-based Data Cleaning tool using Python (with `pandas` and `argparse`), which lets you pass raw CSV or Excel files directly from your terminal and output a cleaned version.

Since you already work with Python and Java, Python is the fastest route here because `pandas` handles missing values, duplicates, and formatting in just a few lines of code.

### 1. Project Setup

Install the required libraries in your terminal:

```bash
pip install pandas openpyxl

```

*(Note: `openpyxl` is required if you plan to pass `.xlsx` Excel files).*

---

### 2. The Console Tool (`cleaner.py`)

Create a file named `cleaner.py` and add the following code. It uses Python's built-in `argparse` to accept command-line flags so you can customize how the data gets cleaned on the fly.

```python
import argparse
import os
import sys
import pandas as pd


def load_data(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        return pd.read_csv(file_path)
    elif ext in ['.xls', '.xlsx']:
        return pd.read_excel(file_path)
    else:
        print('Error: Unsupported file format. Please provide a CSV or Excel file.')
        sys.exit(1)


def save_data(df, output_path):
    ext = os.path.splitext(output_path)[1].lower()
    if ext == '.csv':
        df.to_csv(output_path, index=False)
    elif ext in ['.xls', '.xlsx']:
        df.to_excel(output_path, index=False)
    else:
        print('Error: Output format must be CSV or Excel.')
        sys.exit(1)


def clean_dataframe(df, args):
    initial_rows = len(df)
    print(f'--- Starting Cleaning Process (Initial rows: {initial_rows}) ---')

    # 1. Strip whitespace from column headers and lowercase them
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    print('[✓] Standardized column names.')

    # 2. Trim leading/trailing whitespace from all string cells
    str_cols = df.select_dtypes(include=['object', 'string']).columns
    for col in str_cols:
        df[col] = df[col].apply(
            lambda x: x.strip() if isinstance(x, str) else x
        )
    print('[✓] Trimmed extra whitespace from text fields.')

    # 3. Remove completely empty rows and columns
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)

    # 4. Handle duplicates
    if args.drop_duplicates:
        before = len(df)
        df.drop_duplicates(inplace=True)
        print(f'[✓] Removed {before - len(df)} duplicate rows.')

    # 5. Handle missing values (Nulls / NaNs)
    if args.fill_na is not None:
        df.fillna(args.fill_na, inplace=True)
        print(f"[✓] Filled missing values with '{args.fill_na}'.")
    elif args.drop_na:
        before = len(df)
        df.dropna(inplace=True)
        print(f'[✓] Dropped {before - len(df)} rows containing missing values.')

    print(f'--- Cleaning Complete (Final rows: {len(df)}) ---')
    return df


def main():
    parser = argparse.ArgumentParser(
        description='CLI Data Cleaning Tool for CSV and Excel files.'
    )

    # Required arguments
    parser.add_argument('input', help='Path to the raw input file (CSV/Excel)')
    parser.add_argument(
        '-o',
        '--output',
        default='cleaned_data.csv',
        help='Path for the cleaned output file (default: cleaned_data.csv)',
    )

    # Optional cleaning flags
    parser.add_argument(
        '--drop-duplicates', action='store_true', help='Remove duplicate rows'
    )
    parser.add_argument(
        '--drop-na',
        action='store_true',
        help='Drop rows that have any missing values',
    )
    parser.add_argument(
        '--fill-na',
        type=str,
        default=None,
        help="Fill missing values with a specific string/number (e.g., 'N/A' or '0')",
    )

    args = parser.parse_args()

    df = load_data(args.input)
    cleaned_df = clean_dataframe(df, args)
    save_data(cleaned_df, args.output)
    print(f'Saved cleaned data to: {args.output}')


if __name__ == '__main__':
    main()

```

---

### 3. How to Use It from the Console

**Basic cleaning (standardizes headers, trims whitespace, removes completely blank rows):**

```bash
python cleaner.py raw_sales.csv -o clean_sales.csv

```

**Remove duplicates and drop any row with missing data:**

```bash
python cleaner.py raw_sales.csv --drop-duplicates --drop-na -o clean_sales.csv

```

**Remove duplicates and replace missing cells with "Unknown":**

```bash
python cleaner.py raw_data.xlsx --drop-duplicates --fill-na "Unknown" -o clean_data.xlsx

```

**View all available commands:**

```bash
python cleaner.py --help

```

### Next Steps to Expand This Tool

* **Interactive Mode:** If you don't want to pass flags every time, you can add a `--interactive` flag that prints out missing value counts per column and asks `input()` prompts (e.g., *"Column 'Age' has 5 nulls. Fill with mean, 0, or drop?"*).
* **Type Casting & Regex:** Add rules to automatically strip currency symbols (`₹`, `$`) from numeric columns or standardize date formats using `pd.to_datetime()`.
* **Java Alternative:** If you prefer building this in Java instead (for a standalone `.jar` executable), you can achieve the same architecture using **Picocli** (for CLI arguments) and **OpenCSV** / **Apache POI** (for file parsing).

Do you want to add interactive prompts, custom column rules (like date/currency formatting), or would you prefer to build this in Java?