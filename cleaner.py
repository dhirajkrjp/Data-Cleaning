import argparse
import csv
from datetime import datetime
import json
import os
import re
import sys
import time
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Enable ANSI color support in Windows CMD/PowerShell
os.system('')

# ANSI Console Colors
GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
CYAN = '\033[96m'
BOLD = '\033[1m'
RESET = '\033[0m'

DEFAULT_RULES = {
    'required_columns': ['order_id', 'customer_name', 'total_amount'],
    'email_keywords': ['email'],
    'phone_keywords': ['phone', 'mobile', 'contact'],
    'currency_keywords': ['price', 'amount', 'cost', 'salary', 'revenue'],
    'quantity_keywords': ['quantity', 'qty', 'count'],
    'date_keywords': ['date', 'dob', 'created_at'],
    'name_keywords': ['customer_name', 'full_name', 'name'],
    'city_keywords': ['city', 'location'],
    'status_keywords': ['status', 'category', 'payment_method'],
    'city_aliases': {
        'bangalore': 'Bengaluru',
        'bombay': 'Mumbai',
        'calcutta': 'Kolkata',
        'madras': 'Chennai',
        'new delhi': 'Delhi',
    },
    'invalid_placeholders': [
        'unknown',
        'null',
        'nan',
        'n/a',
        'none',
        'error',
        '-',
    ],
}


def load_or_create_rules(rules_path='rules.json'):
    if not os.path.exists(rules_path):
        with open(rules_path, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_RULES, f, indent=4)
        print(
            f'{CYAN}[INFO] Created default configuration file: {rules_path}{RESET}'
        )
        return DEFAULT_RULES

    with open(rules_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def is_empty_value(val):
    return val is None or (isinstance(val, str) and val.strip() == '')


def parse_number(val):
    """Strips currency symbols/commas and parses to int or float."""
    if isinstance(val, (int, float)):
        return val
    if is_empty_value(val):
        return None
    cleaned = re.sub(r'[₹$,]|Rs\.?', '', str(val), flags=re.IGNORECASE).strip()
    try:
        num = float(cleaned)
        return int(num) if num.is_integer() else round(num, 2)
    except ValueError:
        return None


def standardize_date(val):
    """Converts mixed date formats or Excel datetime objects to YYYY-MM-DD."""
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d'), True
    if is_empty_value(val):
        return None, False

    val_str = str(val).strip()
    date_formats = [
        '%Y-%m-%d',
        '%d/%m/%Y',
        '%d-%b-%Y',
        '%m/%d/%Y',
        '%d-%m-%Y',
        '%Y/%m/%d',
        '%d %b %Y',
    ]
    for fmt in date_formats:
        try:
            dt = datetime.strptime(val_str, fmt)
            formatted = dt.strftime('%Y-%m-%d')
            return formatted, (formatted != val_str)
        except ValueError:
            continue
    return val_str, False


def standardize_phone(val):
    """Normalizes 10-digit or +91 numbers to +91-XXXXXXXXXX."""
    if is_empty_value(val):
        return None, False
    val_str = str(val).strip()
    digits = re.sub(r'\D', '', val_str)
    if len(digits) == 10 and digits[0] in '6789':
        formatted = f'+91-{digits}'
        return formatted, (formatted != val_str)
    elif len(digits) == 12 and digits.startswith('91') and digits[2] in '6789':
        formatted = f'+91-{digits[2:]}'
        return formatted, (formatted != val_str)
    return val_str, False


def load_data(file_path):
    if not os.path.exists(file_path):
        print(f"{RED}Error: File '{file_path}' not found.{RESET}")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    rows = []

    if ext == '.csv':
        with open(file_path, mode='r', encoding='utf-8-sig', newline='') as f:
            rows = list(csv.reader(f))
    elif ext in ['.xls', '.xlsx']:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
    else:
        print(
            f'{RED}Error: Unsupported file format ({ext}). Use CSV or Excel.{RESET}'
        )
        sys.exit(1)

    if not rows:
        print(f'{RED}Error: The input file is empty.{RESET}')
        sys.exit(1)

    return rows[0], rows[1:]


def save_styled_excel(
    headers,
    data,
    output_path,
    sheet_name='Cleaned_Data',
    repaired_coords=None,
    is_rejected=False,
):
    """Saves output to CSV or styled Excel with highlighted repaired cells."""
    ext = os.path.splitext(output_path)[1].lower()
    repaired_coords = repaired_coords or set()

    if ext == '.csv':
        with open(output_path, mode='w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(data)
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_color = '8B0000' if is_rejected else '1F4E78'
    header_fill = PatternFill(
        start_color=header_color, end_color=header_color, fill_type='solid'
    )
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    repaired_fill = PatternFill(
        start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'
    )
    zebra_fill = PatternFill(
        start_color='F7F9FC', end_color='F7F9FC', fill_type='solid'
    )
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9'),
    )

    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='left', vertical='center')
        cell.border = thin_border

    for r_idx, row in enumerate(data, start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name='Calibri', size=10)
            cell.border = thin_border
            if (r_idx - 2, c_idx - 1) in repaired_coords:
                cell.fill = repaired_fill
            elif r_idx % 2 == 0:
                cell.fill = zebra_fill

    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            val_len = len(str(cell.value)) if cell.value is not None else 0
            if val_len > max_len:
                max_len = val_len
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(output_path)


def format_row_preview(headers, row, max_cols=4):
    pairs = []
    for h, val in list(zip(headers, row))[:max_cols]:
        display_val = 'EMPTY' if is_empty_value(val) else str(val)
        pairs.append(f'{h}={display_val}')
    return ' | '.join(pairs)


def write_professional_log(
    log_path,
    input_path,
    output_path,
    quarantine_path,
    args,
    stats,
    initial_rows,
    final_clean_count,
    rejected_count,
    rejected_entries,
    repaired_entries,
    accepted_entries,
    duration_sec,
):
    """Generates a structured, readable audit log file inside the Logs folder."""
    run_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sep_heavy = '=' * 90
    sep_light = '-' * 90

    def pct(count):
        return f'{(count / initial_rows * 100):.1f}%' if initial_rows else '0.0%'

    lines = [
        sep_heavy,
        'ENTERPRISE DATA CLEANING PIPELINE - EXECUTION AUDIT LOG',
        sep_heavy,
        f'Execution Timestamp   : {run_timestamp}',
        f'Execution Duration    : {duration_sec:.2f} seconds',
        f'Source Input File     : {os.path.abspath(input_path)}',
        f'Cleaned Output File   : {os.path.abspath(output_path)}',
        f'Quarantine Output File: {os.path.abspath(quarantine_path) if rejected_count else "N/A (No rows rejected)"}',
        f'Active Pipeline Flags : auto_fix={args.auto_fix} | drop_duplicates={args.drop_duplicates} | drop_na={args.drop_na} | fill_na={args.fill_na}',
        sep_heavy,
        '',
        '1. EXECUTIVE SUMMARY',
        sep_light,
        f"{'Metric':<32} | {'Row Count':<12} | {'Share (%)':<10}",
        sep_light,
        f"{'Total Raw Rows Processed':<32} | {initial_rows:<12} | {'100.0%':<10}",
        f"{'Clean Accepted (Unmodified)':<32} | {stats['accepted']:<12} | {pct(stats['accepted']):<10}",
        f"{'Auto-Repaired & Accepted':<32} | {stats['repaired']:<12} | {pct(stats['repaired']):<10}",
        f"{'Removed - Blank Rows':<32} | {stats['removed_blank']:<12} | {pct(stats['removed_blank']):<10}",
        f"{'Removed - Duplicate Rows':<32} | {stats['removed_duplicate']:<12} | {pct(stats['removed_duplicate']):<10}",
        f"{'Removed - Invalid/Missing Data':<32} | {stats['removed_invalid']:<12} | {pct(stats['removed_invalid']):<10}",
        sep_light,
        f"{'FINAL SAVED CLEAN ROWS':<32} | {final_clean_count:<12} | {pct(final_clean_count):<10}",
        f"{'TOTAL QUARANTINED ROWS':<32} | {rejected_count:<12} | {pct(rejected_count):<10}",
        sep_heavy,
        '',
        f'2. REJECTED & QUARANTINED RECORDS ({len(rejected_entries)} rows)',
        sep_light,
    ]

    if rejected_entries:
        for entry in rejected_entries:
            lines.append(
                f"[Row {entry['row']:03d}] [{entry['category']}] {entry['preview']}"
            )
            lines.append(f"          └── Reason: {entry['reason']}")
    else:
        lines.append('No rows were rejected during this run.')

    lines.extend([
        '',
        sep_heavy,
        f'3. AUTO-REPAIRED & ACCEPTED RECORDS ({len(repaired_entries)} rows)',
        sep_light,
    ])

    if repaired_entries:
        for entry in repaired_entries:
            lines.append(f"[Row {entry['row']:03d}] [REPAIRED] {entry['preview']}")
            for fix in entry['fixes']:
                lines.append(f'          └── Fixed: {fix}')
    else:
        lines.append('No rows required auto-repair during this run.')

    lines.extend([
        '',
        sep_heavy,
        f'4. CLEAN ACCEPTED RECORDS ({len(accepted_entries)} rows)',
        sep_light,
    ])

    if accepted_entries:
        for entry in accepted_entries:
            lines.append(f"[Row {entry['row']:03d}] [ACCEPTED] {entry['preview']}")
    else:
        lines.append('No unmodified clean rows in this dataset.')

    lines.extend(['', sep_heavy, 'END OF AUDIT LOG', sep_heavy])

    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))


def clean_dataset(headers, data, args, rules):
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

    col_index_map = {name: idx for idx, name in enumerate(cleaned_headers)}
    cleaned_rows = []
    rejected_rows = []
    repaired_coords = set()
    seen = set()

    # Structured log buckets
    rejected_entries = []
    repaired_entries = []
    accepted_entries = []

    stats = {
        'accepted': 0,
        'repaired': 0,
        'removed_blank': 0,
        'removed_duplicate': 0,
        'removed_invalid': 0,
    }

    for excel_row_num, row in enumerate(data, start=2):
        current_row = list(row[: len(cleaned_headers)])
        while len(current_row) < len(cleaned_headers):
            current_row.append(None)

        # Check 1: Completely Blank Row
        if all(is_empty_value(c) for c in current_row):
            stats['removed_blank'] += 1
            print(
                f'{RED}[Row {excel_row_num:03d}] REMOVED (BLANK ROW)   -> All cells are empty{RESET}'
            )
            rejected_rows.append(current_row + ['Blank Row'])
            rejected_entries.append({
                'row': excel_row_num,
                'category': 'BLANK ROW',
                'preview': 'All cells are empty',
                'reason': 'Empty record',
            })
            continue

        # Trim whitespace
        current_row = [
            c.strip() if isinstance(c, str) else c for c in current_row
        ]
        preview = format_row_preview(cleaned_headers, current_row)

        # Check 2: Duplicate Row
        row_key = tuple('' if is_empty_value(c) else str(c) for c in current_row)
        if args.drop_duplicates and row_key in seen:
            stats['removed_duplicate'] += 1
            print(
                f'{RED}[Row {excel_row_num:03d}] REMOVED (DUPLICATE)   -> {preview}{RESET}'
            )
            rejected_rows.append(current_row + ['Duplicate Row'])
            rejected_entries.append({
                'row': excel_row_num,
                'category': 'DUPLICATE',
                'preview': preview,
                'reason': 'Exact duplicate of an earlier row',
            })
            continue
        seen.add(row_key)

        row_repairs = []
        row_repaired_cols = set()

        # Check 3: Auto-Fix / Standardization Engine
        if args.auto_fix:
            for c_idx, col_name in enumerate(cleaned_headers):
                val = current_row[c_idx]
                if is_empty_value(val):
                    continue

                # Currency & Quantity parsing (only log if actual value/type changed)
                if any(
                    k in col_name
                    for k in rules['currency_keywords']
                    + rules['quantity_keywords']
                ):
                    parsed = parse_number(val)
                    if parsed is not None and (
                        parsed != val or type(parsed) is not type(val)
                    ):
                        current_row[c_idx] = parsed
                        row_repairs.append(f'{col_name}: {val} -> {parsed}')
                        row_repaired_cols.add(c_idx)

                # Date standardization
                elif any(k in col_name for k in rules['date_keywords']):
                    std_date, changed = standardize_date(val)
                    if changed:
                        current_row[c_idx] = std_date
                        row_repairs.append(f'{col_name}: {val} -> {std_date}')
                        row_repaired_cols.add(c_idx)

                # Phone standardization
                elif any(k in col_name for k in rules['phone_keywords']):
                    std_phone, changed = standardize_phone(val)
                    if changed:
                        current_row[c_idx] = std_phone
                        row_repairs.append(f'{col_name}: {val} -> {std_phone}')
                        row_repaired_cols.add(c_idx)

                # Email lowercase
                elif any(k in col_name for k in rules['email_keywords']):
                    lower_email = str(val).lower()
                    if lower_email != str(val):
                        current_row[c_idx] = lower_email
                        row_repairs.append(
                            f'{col_name}: {val} -> {lower_email}'
                        )
                        row_repaired_cols.add(c_idx)

                # Name Title Case
                elif (
                    any(k in col_name for k in rules['name_keywords'])
                    and 'product' not in col_name
                ):
                    title_name = ' '.join(str(val).split()).title()
                    if title_name != str(val):
                        current_row[c_idx] = title_name
                        row_repairs.append(f'{col_name}: {val} -> {title_name}')
                        row_repaired_cols.add(c_idx)

                # City Alias & Title Case
                elif any(k in col_name for k in rules['city_keywords']):
                    city_clean = str(val).strip()
                    mapped = rules['city_aliases'].get(
                        city_clean.lower(), city_clean.title()
                    )
                    if mapped != str(val):
                        current_row[c_idx] = mapped
                        row_repairs.append(f'{col_name}: {val} -> {mapped}')
                        row_repaired_cols.add(c_idx)

                # Status/Category Title Case
                elif any(k in col_name for k in rules['status_keywords']):
                    if str(val).lower() not in ['upi']:
                        title_stat = str(val).title()
                        if title_stat != str(val):
                            current_row[c_idx] = title_stat
                            row_repairs.append(
                                f'{col_name}: {val} -> {title_stat}'
                            )
                            row_repaired_cols.add(c_idx)

            # Recalculate missing Total Amount if Quantity and Unit Price exist
            qty_idx = next(
                (i for h, i in col_index_map.items() if 'quantity' in h), None
            )
            price_idx = next(
                (i for h, i in col_index_map.items() if 'unit_price' in h), None
            )
            total_idx = next(
                (i for h, i in col_index_map.items() if 'total_amount' in h),
                None,
            )

            if (
                qty_idx is not None
                and price_idx is not None
                and total_idx is not None
            ):
                q_val = parse_number(current_row[qty_idx])
                p_val = parse_number(current_row[price_idx])
                t_val = current_row[total_idx]
                if (
                    is_empty_value(t_val)
                    and isinstance(q_val, (int, float))
                    and isinstance(p_val, (int, float))
                ):
                    calc_total = round(q_val * p_val, 2)
                    if calc_total.is_integer():
                        calc_total = int(calc_total)
                    current_row[total_idx] = calc_total
                    row_repairs.append(
                        f'total_amount: EMPTY -> {calc_total} (Calculated)'
                    )
                    row_repaired_cols.add(total_idx)

        # Check 4: Validate Remaining Issues
        issues = []
        for c_idx, (col_name, cell) in enumerate(
            zip(cleaned_headers, current_row)
        ):
            if is_empty_value(cell):
                issues.append(f'{col_name}: Missing')
            else:
                val_str = str(cell).strip()
                if val_str.lower() in rules['invalid_placeholders']:
                    issues.append(f"{col_name}: Invalid ('{val_str}')")
                elif (
                    any(k in col_name for k in rules['email_keywords'])
                    and '@' not in val_str
                ):
                    issues.append(f"{col_name}: Invalid email ('{val_str}')")
                elif any(
                    k in col_name
                    for k in rules['currency_keywords']
                    + rules['quantity_keywords']
                ):
                    num = parse_number(cell)
                    if num is None or num < 0:
                        issues.append(
                            f"{col_name}: Invalid number ('{val_str}')"
                        )

        updated_preview = format_row_preview(cleaned_headers, current_row)

        if issues:
            issue_str = ', '.join(issues)
            if args.drop_na:
                stats['removed_invalid'] += 1
                print(
                    f'{RED}[Row {excel_row_num:03d}] REMOVED (BAD/MISSING) -> {updated_preview}{RESET}'
                )
                print(f'{RED}          └── Reason: {issue_str}{RESET}')
                rejected_rows.append(current_row + [issue_str])
                rejected_entries.append({
                    'row': excel_row_num,
                    'category': 'BAD/MISSING',
                    'preview': updated_preview,
                    'reason': issue_str,
                })
                continue
            elif args.fill_na is not None:
                for c_idx, cell in enumerate(current_row):
                    if is_empty_value(cell):
                        current_row[c_idx] = args.fill_na
                        row_repaired_cols.add(c_idx)
                        row_repairs.append(
                            f"{cleaned_headers[c_idx]}: EMPTY -> '{args.fill_na}'"
                        )
                dest_row_idx = len(cleaned_rows)
                for c_idx in row_repaired_cols:
                    repaired_coords.add((dest_row_idx, c_idx))
                cleaned_rows.append(current_row)
                stats['repaired'] += 1
                print(
                    f"{YELLOW}[Row {excel_row_num:03d}] FIXED (FILLED '{args.fill_na}') -> {updated_preview}{RESET}"
                )
                print(
                    f'{YELLOW}          └── Issues Handled: {issue_str}{RESET}'
                )
                repaired_entries.append({
                    'row': excel_row_num,
                    'preview': updated_preview,
                    'fixes': row_repairs,
                })
                continue
            else:
                dest_row_idx = len(cleaned_rows)
                for c_idx in row_repaired_cols:
                    repaired_coords.add((dest_row_idx, c_idx))
                cleaned_rows.append(current_row)
                stats['repaired'] += 1
                print(
                    f'{RED}[Row {excel_row_num:03d}] WARNING (MISSING/WRONG) -> {updated_preview}{RESET}'
                )
                print(
                    f'{RED}          └── Remaining Issues: {issue_str}{RESET}'
                )
                repaired_entries.append({
                    'row': excel_row_num,
                    'preview': updated_preview,
                    'fixes': row_repairs + [f'FLAGGED: {issue_str}'],
                })
                continue

        # Clean or Auto-Repaired Row
        dest_row_idx = len(cleaned_rows)
        for c_idx in row_repaired_cols:
            repaired_coords.add((dest_row_idx, c_idx))
        cleaned_rows.append(current_row)

        if row_repairs:
            stats['repaired'] += 1
            print(
                f'{YELLOW}[Row {excel_row_num:03d}] INFO: REPAIRED & ACCEPTED -> {updated_preview}{RESET}'
            )
            print(
                f"{YELLOW}          └── Auto-Fixed: {', '.join(row_repairs)}{RESET}"
            )
            repaired_entries.append({
                'row': excel_row_num,
                'preview': updated_preview,
                'fixes': row_repairs,
            })
        else:
            stats['accepted'] += 1
            print(
                f'{GREEN}[Row {excel_row_num:03d}] INFO: ACCEPTED            -> {updated_preview}{RESET}'
            )
            accepted_entries.append({
                'row': excel_row_num,
                'preview': updated_preview,
            })

    # Console Summary Report
    print(f'\n{CYAN}{BOLD}=== CLEANING SUMMARY REPORT ==={RESET}')
    print(f'Total Raw Rows Processed  : {initial_rows}')
    print(f"{GREEN}✔ Clean Accepted Rows     : {stats['accepted']}{RESET}")
    print(f"{YELLOW}✔ Auto-Repaired / Flagged : {stats['repaired']}{RESET}")
    print(f"{RED}✖ Removed Blank Rows      : {stats['removed_blank']}{RESET}")
    print(f"{RED}✖ Removed Duplicates      : {stats['removed_duplicate']}{RESET}")
    print(f"{RED}✖ Removed Invalid Rows    : {stats['removed_invalid']}{RESET}")
    print(f'{BOLD}Final Saved Clean Rows    : {len(cleaned_rows)}{RESET}')
    print(f'{BOLD}Total Quarantined Rows    : {len(rejected_rows)}{RESET}\n')

    return (
        cleaned_headers,
        cleaned_rows,
        rejected_rows,
        repaired_coords,
        stats,
        rejected_entries,
        repaired_entries,
        accepted_entries,
    )


def run_interactive_menu(args):
    print(f'\n{CYAN}{BOLD}=== INTERACTIVE DATA CLEANING WIZARD ==={RESET}')
    inp = input(f'Enter input file path [{args.input}]: ').strip()
    if inp:
        args.input = inp

    out = input(f'Enter cleaned output file path [{args.output}]: ').strip()
    if out:
        args.output = out

    args.drop_duplicates = (
        input('Remove duplicate rows? (y/n) [y]: ').strip().lower() != 'n'
    )
    args.auto_fix = (
        input(
            'Auto-fix currencies, dates, phones, cities & totals? (y/n) [y]: '
        )
        .strip()
        .lower()
        != 'n'
    )

    na_choice = input(
        'How to handle missing values? [1] Keep & Flag  [2] Drop & Quarantine  [3] Fill with text (default 1): '
    ).strip()
    if na_choice == '2':
        args.drop_na = True
    elif na_choice == '3':
        fill_val = (
            input('Enter replacement value [Unknown]: ').strip() or 'Unknown'
        )
        args.fill_na = fill_val

    quarantine = input(
        f'Save rejected/removed rows to file [{args.quarantine}]: '
    ).strip()
    if quarantine:
        args.quarantine = quarantine


def process_single_file(input_path, output_path, quarantine_path, args, rules):
    start_time = time.time()
    headers, data = load_data(input_path)
    (
        cleaned_headers,
        cleaned_data,
        rejected_data,
        repaired_coords,
        stats,
        rejected_entries,
        repaired_entries,
        accepted_entries,
    ) = clean_dataset(headers, data, args, rules)

    save_styled_excel(
        cleaned_headers,
        cleaned_data,
        output_path,
        sheet_name='Cleaned_Data',
        repaired_coords=repaired_coords,
    )
    print(f'{GREEN}[SAVED] Cleaned dataset written to: {output_path}{RESET}')

    if rejected_data and quarantine_path:
        rejected_headers = cleaned_headers + ['rejection_reason']
        save_styled_excel(
            rejected_headers,
            rejected_data,
            quarantine_path,
            sheet_name='Rejected_Data',
            is_rejected=True,
        )
        print(
            f'{RED}[QUARANTINE] {len(rejected_data)} removed rows saved to: {quarantine_path}{RESET}'
        )

    duration_sec = time.time() - start_time

    # Auto-generate timestamped log inside ./Logs folder unless custom path given
    if args.log_file:
        log_path = args.log_file
    else:
        base_dir = os.path.dirname(os.path.abspath(input_path))
        logs_dir = os.path.join(base_dir, 'Logs')
        input_stem = os.path.splitext(os.path.basename(input_path))[0]
        ts = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        log_path = os.path.join(logs_dir, f'audit_{input_stem}_{ts}.log')

    write_professional_log(
        log_path=log_path,
        input_path=input_path,
        output_path=output_path,
        quarantine_path=quarantine_path,
        args=args,
        stats=stats,
        initial_rows=len(data),
        final_clean_count=len(cleaned_data),
        rejected_count=len(rejected_data),
        rejected_entries=rejected_entries,
        repaired_entries=repaired_entries,
        accepted_entries=accepted_entries,
        duration_sec=duration_sec,
    )
    print(f'{CYAN}[LOG] Structured audit report saved to: {log_path}{RESET}')


def main():
    parser = argparse.ArgumentParser(
        description='Enterprise CLI Data Cleaning Tool (Pure Python).'
    )

    parser.add_argument(
        'input',
        nargs='?',
        default='raw_data.xlsx',
        help='Input CSV/Excel file (default: raw_data.xlsx)',
    )
    parser.add_argument(
        '-o',
        '--output',
        default='cleaned_data.xlsx',
        help='Cleaned output file (default: cleaned_data.xlsx)',
    )
    parser.add_argument(
        '-q',
        '--quarantine',
        default='rejected_data.xlsx',
        help='Output file for removed/rejected rows (default: rejected_data.xlsx)',
    )
    parser.add_argument(
        '--drop-duplicates', action='store_true', help='Remove duplicate rows'
    )
    parser.add_argument(
        '--auto-fix',
        action='store_true',
        help='Auto-repair currencies, dates, phones, cities, and totals',
    )
    parser.add_argument(
        '--drop-na',
        action='store_true',
        help='Drop rows with missing/invalid values into quarantine file',
    )
    parser.add_argument(
        '--fill-na',
        type=str,
        default=None,
        help="Fill missing cells with a value (e.g., 'Unknown')",
    )
    parser.add_argument(
        '--log-file',
        type=str,
        default=None,
        help='Custom path for audit log (default: auto-saved inside ./Logs/)',
    )
    parser.add_argument(
        '--rules',
        type=str,
        default='rules.json',
        help='Path to JSON rules configuration file',
    )
    parser.add_argument(
        '--folder',
        type=str,
        default=None,
        help='Batch clean all CSV/Excel files in a directory',
    )
    parser.add_argument(
        '-i',
        '--interactive',
        action='store_true',
        help='Launch interactive setup wizard',
    )

    args = parser.parse_args()
    rules = load_or_create_rules(args.rules)

    if args.interactive:
        run_interactive_menu(args)

    if args.folder:
        if not os.path.isdir(args.folder):
            print(f"{RED}Error: Folder '{args.folder}' does not exist.{RESET}")
            sys.exit(1)
        out_dir = os.path.join(args.folder, 'cleaned_output')
        os.makedirs(out_dir, exist_ok=True)
        files = [
            f
            for f in os.listdir(args.folder)
            if f.lower().endswith(('.csv', '.xlsx', '.xls'))
        ]
        if not files:
            print(f"{YELLOW}No CSV or Excel files found in '{args.folder}'.{RESET}")
            return
        for fname in files:
            in_p = os.path.join(args.folder, fname)
            out_p = os.path.join(out_dir, f'cleaned_{fname}')
            rej_p = os.path.join(out_dir, f'rejected_{fname}')
            print(f'\n{CYAN}{BOLD}>>> Processing Batch File: {fname}{RESET}')
            process_single_file(in_p, out_p, rej_p, args, rules)
    else:
        process_single_file(
            args.input, args.output, args.quarantine, args, rules
        )


if __name__ == '__main__':
    main()