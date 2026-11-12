import argparse
import csv
import os
import sys
import openpyxl


def load_data(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    rows = []

    if ext == '.csv':
        with open(file_path, mode='r', encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
    elif ext in ['.xls', '.xlsx']:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
    else:
        print('Error: Unsupported file format. Please provide a CSV or Excel file.')
        sys.exit(1)

    if not rows:
        print('Error: The input file is empty.')
        sys.exit(1)

    headers = rows[0]
    data = rows[1:]
    return headers, data


def save_data(headers, data, output_path):
    ext = os.path.splitext(output_path)[1].lower()

    if ext == '.csv':
        with open(output_path, mode='w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(data)
    elif ext in ['.xls', '.xlsx']:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Cleaned_Data'
        ws.append(headers)
        for row in data:
            ws.append(row)
        wb.save(output_path)
    else:
        print('Error: Output format must be CSV or Excel.')
        sys.exit(1)


def is_empty_value(val):
    return val is None or (isinstance(val, str) and val.strip() == '')


def clean_dataset(headers, data, args):
    initial_rows = len(data)
    print(f'--- Starting Cleaning Process (Initial rows: {initial_rows}) ---')

    # 1. Clean headers (strip whitespace, lowercase, replace spaces with underscores)
    cleaned_headers = []
    for idx, h in enumerate(headers):
        if is_empty_value(h):
            cleaned_headers.append(f'column_{idx + 1}')
        else:
            cleaned_headers.append(str(h).strip().lower().replace(' ', '_'))
    print('[✓] Standardized column names.')

    # 2. Trim whitespace from text cells and drop completely blank rows
    cleaned_rows = []
    blank_rows_removed = 0

    for row in data:
        # Pad or trim row to match header length
        normalized_row = list(row[: len(cleaned_headers)])
        while len(normalized_row) < len(cleaned_headers):
            normalized_row.append(None)

        trimmed_row = [
            cell.strip() if isinstance(cell, str) else cell
            for cell in normalized_row
        ]

        # Skip completely empty rows
        if all(is_empty_value(cell) for cell in trimmed_row):
            blank_rows_removed += 1
            continue

        cleaned_rows.append(trimmed_row)

    print(
        f'[✓] Trimmed text whitespace and removed {blank_rows_removed} completely blank rows.'
    )

    # 3. Handle duplicates
    if args.drop_duplicates:
        seen = set()
        unique_rows = []
        dup_count = 0
        for row in cleaned_rows:
            # Normalize None and empty strings for accurate duplicate detection
            row_key = tuple('' if is_empty_value(c) else str(c) for c in row)
            if row_key not in seen:
                seen.add(row_key)
                unique_rows.append(row)
            else:
                dup_count += 1
        cleaned_rows = unique_rows
        print(f'[✓] Removed {dup_count} duplicate rows.')

    # 4. Handle missing values
    if args.fill_na is not None:
        for r_idx, row in enumerate(cleaned_rows):
            cleaned_rows[r_idx] = [
                args.fill_na if is_empty_value(cell) else cell for cell in row
            ]
        print(f"[✓] Filled missing values with '{args.fill_na}'.")
    elif args.drop_na:
        before = len(cleaned_rows)
        cleaned_rows = [
            row
            for row in cleaned_rows
            if not any(is_empty_value(cell) for cell in row)
        ]
        print(f'[✓] Dropped {before - len(cleaned_rows)} rows containing missing values.')

    print(f'--- Cleaning Complete (Final rows: {len(cleaned_rows)}) ---')
    return cleaned_headers, cleaned_rows


def main():
    parser = argparse.ArgumentParser(
        description='CLI Data Cleaning Tool for CSV and Excel files (Pure Python).'
    )

    parser.add_argument('input', help='Path to the raw input file (CSV/Excel)')
    parser.add_argument(
        '-o',
        '--output',
        default='cleaned_data.xlsx',
        help='Path for the cleaned output file (default: cleaned_data.xlsx)',
    )
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
        help="Fill missing values with a specific string/number (e.g., 'Unknown' or '0')",
    )

    args = parser.parse_args()

    headers, data = load_data(args.input)
    cleaned_headers, cleaned_data = clean_dataset(headers, data, args)
    save_data(cleaned_headers, cleaned_data, args.output)
    print(f'Saved cleaned data to: {args.output}')


if __name__ == '__main__':
    main()