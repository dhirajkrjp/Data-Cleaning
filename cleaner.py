import argparse
import csv
import os
import sys
import openpyxl

# Enable ANSI color support in Windows CMD/PowerShell
os.system('')

# ANSI Color Codes
GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
CYAN = '\033[96m'
BOLD = '\033[1m'
RESET = '\033[0m'


def load_data(file_path):
    if not os.path.exists(file_path):
        print(f"{RED}Error: File '{file_path}' not found.{RESET}")
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
        print(
            f'{RED}Error: Unsupported file format. Please provide a CSV or Excel file.{RESET}'
        )
        sys.exit(1)

    if not rows:
        print(f'{RED}Error: The input file is empty.{RESET}')
        sys.exit(1)

    return rows[0], rows[1:]


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
        print(f'{RED}Error: Output format must be CSV or Excel.{RESET}')
        sys.exit(1)


def is_empty_value(val):
    return val is None or (isinstance(val, str) and val.strip() == '')


def is_wrong_value(col_name, val):
    """Detects invalid/wrong values based on column context."""
    if is_empty_value(val):
        return True, 'Missing/Empty'

    val_str = str(val).strip()
    if val_str.lower() in ['unknown', 'null', 'nan', 'n/a', 'none', 'error']:
        return True, f"Invalid placeholder ('{val_str}')"

    # Check email validity
    if 'email' in col_name and '@' not in val_str:
        return True, f"Invalid email ('{val_str}')"

    # Check numeric columns for non-numeric characters (e.g., currency symbols or negative numbers)
    if any(k in col_name for k in ['quantity', 'price', 'amount']):
        cleaned_num = (
            val_str.replace('₹', '')
            .replace('Rs.', '')
            .replace(',', '')
            .strip()
        )
        try:
            num = float(cleaned_num)
            if num < 0:
                return True, f'Negative value ({val_str})'
        except ValueError:
            return True, f"Non-numeric value ('{val_str}')"

    return False, ''


def format_row_preview(headers, row, max_cols=5):
    """Creates a compact readable string of the row for console logging."""
    pairs = []
    for h, val in list(zip(headers, row))[:max_cols]:
        display_val = 'EMPTY' if is_empty_value(val) else str(val).strip()
        pairs.append(f'{h}={display_val}')
    return ' | '.join(pairs)


def clean_dataset(headers, data, args):
    initial_rows = len(data)
    print(
        f'\n{CYAN}{BOLD}=== STARTING DATA CLEANING AUDIT (Total Raw Rows: {initial_rows}) ==={RESET}\n'
    )

    # 1. Standardize headers
    cleaned_headers = []
    for idx, h in enumerate(headers):
        if is_empty_value(h):
            cleaned_headers.append(f'column_{idx + 1}')
        else:
            cleaned_headers.append(str(h).strip().lower().replace(' ', '_'))

    cleaned_rows = []
    seen = set()

    stats = {
        'accepted': 0,
        'fixed': 0,
        'removed_blank': 0,
        'removed_duplicate': 0,
        'removed_invalid': 0,
    }

    for excel_row_num, row in enumerate(data, start=2):
        # Normalize row length
        normalized_row = list(row[: len(cleaned_headers)])
        while len(normalized_row) < len(cleaned_headers):
            normalized_row.append(None)

        # Check 1: Completely Blank Row
        if all(is_empty_value(cell) for cell in normalized_row):
            stats['removed_blank'] += 1
            print(
                f'{RED}[Row {excel_row_num:03d}] REMOVED (BLANK ROW)   -> All cells are empty{RESET}'
            )
            continue

        # Trim whitespace from text cells
        trimmed_row = [
            cell.strip() if isinstance(cell, str) else cell
            for cell in normalized_row
        ]
        preview = format_row_preview(cleaned_headers, trimmed_row)

        # Check 2: Duplicate Row
        row_key = tuple('' if is_empty_value(c) else str(c) for c in trimmed_row)
        if args.drop_duplicates and row_key in seen:
            stats['removed_duplicate'] += 1
            print(
                f'{RED}[Row {excel_row_num:03d}] REMOVED (DUPLICATE)   -> {preview}{RESET}'
            )
            continue
        seen.add(row_key)

        # Check 3: Inspect cells for Missing or Wrong Data
        issues = []
        for col_name, cell in zip(cleaned_headers, trimmed_row):
            is_bad, reason = is_wrong_value(col_name, cell)
            if is_bad:
                issues.append(f'{col_name}: {reason}')

        if issues:
            issue_summary = ', '.join(issues)
            if args.drop_na:
                stats['removed_invalid'] += 1
                print(
                    f'{RED}[Row {excel_row_num:03d}] REMOVED (BAD/MISSING) -> {preview}{RESET}'
                )
                print(
                    f'{RED}          └── Reason: {issue_summary}{RESET}'
                )
                continue
            elif args.fill_na is not None:
                # Replace empty cells with fill_na value
                fixed_row = [
                    args.fill_na if is_empty_value(cell) else cell
                    for cell in trimmed_row
                ]
                cleaned_rows.append(fixed_row)
                stats['fixed'] += 1
                print(
                    f"{YELLOW}[Row {excel_row_num:03d}] FIXED (FILLED '{args.fill_na}') -> {preview}{RESET}"
                )
                print(
                    f'{YELLOW}          └── Issues Found: {issue_summary}{RESET}'
                )
                continue
            else:
                # Kept in dataset, but flagged in red as containing missing/wrong data
                cleaned_rows.append(trimmed_row)
                stats['fixed'] += 1
                print(
                    f'{RED}[Row {excel_row_num:03d}] WARNING (MISSING/WRONG DATA) -> {preview}{RESET}'
                )
                print(
                    f'{RED}          └── Issues Found: {issue_summary}{RESET}'
                )
                continue

        # Check 4: Clean & Accepted Row
        cleaned_rows.append(trimmed_row)
        stats['accepted'] += 1
        print(
            f'{GREEN}[Row {excel_row_num:03d}] INFO: ACCEPTED        -> {preview}{RESET}'
        )

    # Final Summary Report
    print(f'\n{CYAN}{BOLD}=== CLEANING SUMMARY REPORT ==={RESET}')
    print(f'Total Raw Rows Processed : {initial_rows}')
    print(f"{GREEN}✔ Accepted Clean Rows    : {stats['accepted']}{RESET}")
    print(f"{YELLOW}✔ Fixed / Flagged Rows   : {stats['fixed']}{RESET}")
    print(f"{RED}✖ Removed Blank Rows     : {stats['removed_blank']}{RESET}")
    print(f"{RED}✖ Removed Duplicates     : {stats['removed_duplicate']}{RESET}")
    print(f"{RED}✖ Removed Invalid Rows   : {stats['removed_invalid']}{RESET}")
    print(f'{BOLD}Final Saved Rows         : {len(cleaned_rows)}{RESET}\n')

    return cleaned_headers, cleaned_rows


def main():
    parser = argparse.ArgumentParser(
        description='CLI Data Cleaning Tool with Row-by-Row Console Audit.'
    )

    parser.add_argument(
        'input',
        nargs='?',
        default='raw_data.xlsx',
        help='Path to the raw input file (default: raw_data.xlsx)',
    )
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
        help='Drop rows that have any missing or invalid values',
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
    print(f'{GREEN}Saved cleaned data to: {args.output}{RESET}')


if __name__ == '__main__':
    main()