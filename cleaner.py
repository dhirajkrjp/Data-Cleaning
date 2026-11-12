import argparse
import ast
import csv
from datetime import datetime
import json
import operator
import os
import re
import statistics
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
MAGENTA = '\033[95m'
BOLD = '\033[1m'
RESET = '\033[0m'

# Safe AST operators for dynamic formula evaluation
SAFE_OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.FloorDiv: operator.floordiv,
    ast.Mod: operator.mod,
}

DOMAIN_PRESETS = {
    'general': {
        'domain': 'general',
        'description': 'Universal domain-agnostic cleaning profile',
        'email_keywords': ['email', 'mail'],
        'phone_keywords': ['phone', 'mobile', 'contact', 'tel'],
        'currency_keywords': [
            'price',
            'amount',
            'cost',
            'salary',
            'revenue',
            'fee',
            'pay',
            'total',
        ],
        'quantity_keywords': ['quantity', 'qty', 'count', 'units', 'age'],
        'date_keywords': ['date', 'dob', 'created_at', 'updated_at', 'timestamp'],
        'name_keywords': ['customer_name', 'employee_name', 'patient_name', 'full_name', 'name'],
        'city_keywords': ['city', 'location', 'branch'],
        'title_case_keywords': ['status', 'category', 'department', 'payment_method'],
        'city_aliases': {
            'bangalore': 'Bengaluru',
            'bombay': 'Mumbai',
            'calcutta': 'Kolkata',
            'madras': 'Chennai',
            'new delhi': 'Delhi',
        },
        'invalid_placeholders': ['unknown', 'null', 'nan', 'n/a', 'none', 'error', '-'],
        'computed_columns': [],
        'regex_validations': {},
        'allowed_values': {},
        'column_overrides': {},
    },
    'ecommerce': {
        'domain': 'ecommerce',
        'description': 'E-Commerce & Retail Order Dataset Profile',
        'email_keywords': ['email'],
        'phone_keywords': ['phone', 'mobile', 'contact'],
        'currency_keywords': ['price', 'amount', 'cost', 'discount', 'revenue'],
        'quantity_keywords': ['quantity', 'qty'],
        'date_keywords': ['date'],
        'name_keywords': ['customer_name', 'full_name', 'name'],
        'city_keywords': ['city'],
        'title_case_keywords': ['status', 'category', 'payment_method'],
        'city_aliases': {
            'bangalore': 'Bengaluru',
            'bombay': 'Mumbai',
            'calcutta': 'Kolkata',
            'madras': 'Chennai',
            'new delhi': 'Delhi',
        },
        'invalid_placeholders': ['unknown', 'null', 'nan', 'n/a', 'none', 'error', '-'],
        'computed_columns': [
            {
                'target': 'total_amount',
                'formula': 'quantity * unit_price_(inr)',
                'only_if_empty': True,
            }
        ],
        'regex_validations': {
            'order_id': r'^ORD-\d+$',
        },
        'allowed_values': {
            'order_status': ['Completed', 'Pending', 'Cancelled', 'Refunded', 'In Transit'],
        },
        'column_overrides': {},
    },
    'hr': {
        'domain': 'hr',
        'description': 'Human Resources & Payroll Dataset Profile',
        'email_keywords': ['email'],
        'phone_keywords': ['phone', 'mobile'],
        'currency_keywords': ['salary', 'pay', 'hra', 'bonus', 'deductions', 'ctc'],
        'quantity_keywords': ['age', 'experience', 'leaves'],
        'date_keywords': ['date', 'dob', 'joining'],
        'name_keywords': ['employee_name', 'full_name', 'name'],
        'city_keywords': ['city', 'location'],
        'title_case_keywords': ['department', 'designation', 'employment_type', 'status'],
        'city_aliases': {
            'bangalore': 'Bengaluru',
            'bombay': 'Mumbai',
            'new delhi': 'Delhi',
        },
        'invalid_placeholders': ['unknown', 'null', 'nan', 'n/a', 'none', 'error', '-'],
        'computed_columns': [
            {
                'target': 'net_salary',
                'formula': 'basic_pay + hra - deductions',
                'only_if_empty': True,
            }
        ],
        'regex_validations': {
            'employee_id': r'^EMP-\d+$',
        },
        'allowed_values': {
            'employment_type': ['Full-Time', 'Part-Time', 'Contract', 'Intern'],
        },
        'column_overrides': {},
    },
    'finance': {
        'domain': 'finance',
        'description': 'Accounting, Invoicing & Tax Compliance Profile',
        'email_keywords': ['email'],
        'phone_keywords': ['phone', 'mobile'],
        'currency_keywords': ['amount', 'price', 'tax', 'gst', 'debit', 'credit', 'balance'],
        'quantity_keywords': ['quantity', 'qty'],
        'date_keywords': ['date', 'due'],
        'name_keywords': ['client_name', 'vendor_name', 'name'],
        'city_keywords': ['city', 'state'],
        'title_case_keywords': ['status', 'payment_mode'],
        'city_aliases': {'bangalore': 'Bengaluru', 'bombay': 'Mumbai'},
        'invalid_placeholders': ['unknown', 'null', 'nan', 'n/a', 'none', 'error', '-'],
        'computed_columns': [
            {
                'target': 'total_amount',
                'formula': 'subtotal + tax_amount',
                'only_if_empty': True,
            }
        ],
        'regex_validations': {
            'pan_number': r'^[A-Z]{5}[0-9]{4}[A-Z]{1}$',
            'gstin': r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$',
        },
        'allowed_values': {
            'invoice_status': ['Paid', 'Unpaid', 'Overdue', 'Cancelled'],
        },
        'column_overrides': {},
    },
    'healthcare': {
        'domain': 'healthcare',
        'description': 'Clinical & Patient Records Profile',
        'email_keywords': ['email'],
        'phone_keywords': ['phone', 'mobile', 'emergency_contact'],
        'currency_keywords': ['fee', 'bill', 'charge', 'amount'],
        'quantity_keywords': ['age', 'weight', 'heart_rate'],
        'date_keywords': ['date', 'dob', 'admission', 'discharge'],
        'name_keywords': ['patient_name', 'doctor_name', 'name'],
        'city_keywords': ['city'],
        'title_case_keywords': ['gender', 'diagnosis', 'department', 'status'],
        'city_aliases': {'bangalore': 'Bengaluru', 'bombay': 'Mumbai'},
        'invalid_placeholders': ['unknown', 'null', 'nan', 'n/a', 'none', 'error', '-'],
        'computed_columns': [],
        'regex_validations': {
            'patient_id': r'^PAT-\d+$',
            'blood_group': r'^(A|B|AB|O)[+-]$',
        },
        'allowed_values': {
            'gender': ['Male', 'Female', 'Other'],
        },
        'column_overrides': {},
    },
}


# ==============================================================================
# 1. PROFILE & RULE MANAGEMENT
# ==============================================================================
def ensure_domain_profiles(profiles_dir='profiles'):
    os.makedirs(profiles_dir, exist_ok=True)
    for domain_name, preset in DOMAIN_PRESETS.items():
        path = os.path.join(profiles_dir, f'{domain_name}.json')
        if not os.path.exists(path):
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(preset, f, indent=4)


def load_rules(args):
    ensure_domain_profiles('profiles')

    if args.profile:
        if not os.path.exists(args.profile):
            print(f"{RED}Error: Profile file '{args.profile}' not found.{RESET}")
            sys.exit(1)
        with open(args.profile, 'r', encoding='utf-8') as f:
            rules = json.load(f)
        print(f'{CYAN}[PROFILE] Loaded custom profile: {args.profile}{RESET}')
        return rules

    domain_file = os.path.join('profiles', f'{args.domain.lower()}.json')
    if os.path.exists(domain_file):
        with open(domain_file, 'r', encoding='utf-8') as f:
            rules = json.load(f)
        print(f'{CYAN}[PROFILE] Active domain preset: {args.domain.upper()} ({domain_file}){RESET}')
        return rules

    return DOMAIN_PRESETS['general']


# ==============================================================================
# 2. DATA PARSING, FORMULA EVALUATION & NORMALIZATION HELPERS
# ==============================================================================
def is_empty_value(val):
    return val is None or (isinstance(val, str) and val.strip() == '')


def parse_number(val):
    """Strips currency symbols/commas and parses to int or float."""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return int(val) if float(val).is_integer() else round(float(val), 2)
    if is_empty_value(val):
        return None
    cleaned = re.sub(r'[₹$€£,]|Rs\.?', '', str(val), flags=re.IGNORECASE).strip()
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
        '%d-%B-%Y',
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


def safe_eval_formula(formula_str, row_dict):
    """Safely evaluates arithmetic formulas using column variables (no eval)."""
    # Replace column tokens with numeric values
    tokens = re.findall(r'[a-zA-Z_][a-zA-Z0-9_()]*', formula_str)
    expr = formula_str
    # Sort longer column names first to prevent partial replacements
    for token in sorted(set(tokens), key=len, reverse=True):
        if token not in row_dict:
            return None
        num_val = parse_number(row_dict[token])
        if num_val is None:
            return None
        expr = expr.replace(token, str(num_val))

    def _eval_node(node):
        if isinstance(node, ast.Expression):
            return _eval_node(node.body)
        elif isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return node.value
        elif isinstance(node, ast.BinOp) and type(node.op) in SAFE_OPERATORS:
            left = _eval_node(node.left)
            right = _eval_node(node.right)
            if type(node.op) in (ast.Div, ast.FloorDiv, ast.Mod) and right == 0:
                raise ZeroDivisionError
            return SAFE_OPERATORS[type(node.op)](left, right)
        elif isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
            return -_eval_node(node.operand)
        else:
            raise ValueError('Unsupported expression')

    try:
        tree = ast.parse(expr, mode='eval')
        res = _eval_node(tree)
        res = round(float(res), 2)
        return int(res) if res.is_integer() else res
    except Exception:
        return None


# ==============================================================================
# 3. FILE I/O (MULTI-SHEET EXCEL, CSV, JSON)
# ==============================================================================
def get_excel_sheets(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True)
    sheets = wb.sheetnames
    wb.close()
    return sheets


def load_data(file_path, sheet_name=None):
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
        ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
    elif ext == '.json':
        with open(file_path, mode='r', encoding='utf-8') as f:
            records = json.load(f)
        if not records or not isinstance(records, list):
            print(f'{RED}Error: JSON file must contain a non-empty list of objects.{RESET}')
            sys.exit(1)
        headers = list(records[0].keys())
        data = [[rec.get(h) for h in headers] for rec in records]
        return headers, data
    else:
        print(f'{RED}Error: Unsupported file format ({ext}). Use CSV, Excel, or JSON.{RESET}')
        sys.exit(1)

    if not rows:
        print(f'{RED}Error: The input file/sheet is empty.{RESET}')
        sys.exit(1)

    return rows[0], rows[1:]


def save_output(
    headers,
    data,
    output_path,
    sheet_name='Cleaned_Data',
    repaired_coords=None,
    is_rejected=False,
):
    """Exports data to Styled Excel (.xlsx), CSV (.csv), or JSON (.json)."""
    ext = os.path.splitext(output_path)[1].lower()
    repaired_coords = repaired_coords or set()

    if ext == '.csv':
        with open(output_path, mode='w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(data)
        return

    if ext == '.json':
        records = [dict(zip(headers, row)) for row in data]
        with open(output_path, mode='w', encoding='utf-8') as f:
            json.dump(records, f, indent=4, default=str)
        return

    # Styled Excel Output
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    header_color = '8B0000' if is_rejected else '1F4E78'
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    repaired_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    zebra_fill = PatternFill(start_color='F7F9FC', end_color='F7F9FC', fill_type='solid')
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
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 42)

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    wb.save(output_path)


# ==============================================================================
# 4. DYNAMIC SCHEMA PROFILER & INTERACTIVE COLUMN MAPPER
# ==============================================================================
def standardize_headers(raw_headers):
    cleaned = []
    for idx, h in enumerate(raw_headers):
        if is_empty_value(h):
            cleaned.append(f'column_{idx + 1}')
        else:
            cleaned.append(str(h).strip().lower().replace(' ', '_'))
    return cleaned


def profile_schema(headers, data):
    """Inspects rows to infer column types, null rates, numeric stats, and samples."""
    std_headers = standardize_headers(headers)
    total = len(data)
    profile = []

    for c_idx, col_name in enumerate(std_headers):
        non_empty = [
            row[c_idx]
            for row in data
            if c_idx < len(row) and not is_empty_value(row[c_idx])
        ]
        null_count = total - len(non_empty)
        null_pct = (null_count / total * 100) if total else 0.0

        # Infer dominant type
        num_hits = sum(1 for v in non_empty if parse_number(v) is not None)
        date_hits = sum(1 for v in non_empty if standardize_date(v)[1] or (isinstance(v, str) and re.match(r'^\d{4}-\d{2}-\d{2}$', v.strip())))
        email_hits = sum(1 for v in non_empty if isinstance(v, str) and '@' in v and '.' in v)
        phone_hits = sum(1 for v in non_empty if len(re.sub(r'\D', '', str(v))) in (10, 12) and any(k in col_name for k in ['phone', 'mobile', 'contact']))

        n = max(len(non_empty), 1)
        if email_hits / n > 0.6:
            inferred_type = 'Email'
        elif phone_hits / n > 0.6:
            inferred_type = 'Phone'
        elif date_hits / n > 0.6:
            inferred_type = 'Date'
        elif num_hits / n > 0.7:
            inferred_type = 'Numeric'
        else:
            inferred_type = 'Text'

        # Compute numeric stats for mean/median imputation & outlier detection
        numeric_vals = [parse_number(v) for v in non_empty if parse_number(v) is not None]
        col_stats = {}
        if len(numeric_vals) >= 2:
            col_stats['mean'] = round(statistics.mean(numeric_vals), 2)
            col_stats['median'] = round(statistics.median(numeric_vals), 2)
            col_stats['stdev'] = round(statistics.stdev(numeric_vals), 2)
            col_stats['min'] = min(numeric_vals)
            col_stats['max'] = max(numeric_vals)

        samples = ', '.join(str(v).strip() for v in non_empty[:3])[:35]
        profile.append({
            'index': c_idx,
            'name': col_name,
            'type': inferred_type,
            'null_count': null_count,
            'null_pct': round(null_pct, 1),
            'stats': col_stats,
            'samples': samples,
        })

    return std_headers, profile


def print_schema_table(profile):
    print(f'\n{CYAN}{BOLD}=== DYNAMIC SCHEMA PROFILER REPORT ==={RESET}')
    print(f"{'#':<3} | {'Column Name':<22} | {'Inferred Type':<13} | {'Missing':<12} | {'Sample Values'}")
    print('-' * 92)
    for col in profile:
        miss_str = f"{col['null_count']} ({col['null_pct']}%)"
        color = RED if col['null_pct'] > 15 else (YELLOW if col['null_count'] > 0 else GREEN)
        print(
            f"{col['index']+1:<3} | {col['name']:<22} | {col['type']:<13} | {color}{miss_str:<12}{RESET} | {col['samples']}"
        )
    print('-' * 92 + '\n')


def run_column_mapper(profile, rules):
    """Interactive per-column rule builder with option to save reusable profile."""
    print(f'{MAGENTA}{BOLD}=== INTERACTIVE COLUMN-BY-COLUMN RULE BUILDER ==={RESET}')
    print('Configure custom cleaning behavior for specific columns (press Enter to keep default):\n')

    overrides = rules.get('column_overrides', {})

    for col in profile:
        cname = col['name']
        ctype = col['type']
        print(f"{BOLD}Column [{col['index']+1}/{len(profile)}]: {CYAN}{cname}{RESET} (Detected: {ctype}, Missing: {col['null_pct']}%)")
        print('  [0] Auto/Default   [1] Required (Drop row if empty)   [2] Fill Missing (Custom/Mean/Median)')
        print('  [3] Title Case     [4] Uppercase                     [5] Regex Validation   [6] Ignore Column')
        choice = input('  Select action [0]: ').strip()

        if choice == '1':
            overrides[cname] = {'required': True}
        elif choice == '2':
            if ctype == 'Numeric' and col['stats']:
                print(f"    Stats -> Mean: {col['stats']['mean']} | Median: {col['stats']['median']}")
                fill_mode = input('    Enter value, or type MEAN / MEDIAN: ').strip()
                if fill_mode.upper() == 'MEAN':
                    overrides[cname] = {'fill_value': col['stats']['mean']}
                elif fill_mode.upper() == 'MEDIAN':
                    overrides[cname] = {'fill_value': col['stats']['median']}
                else:
                    overrides[cname] = {'fill_value': fill_mode}
            else:
                fill_val = input('    Enter replacement value: ').strip()
                overrides[cname] = {'fill_value': fill_val}
        elif choice == '3':
            overrides[cname] = {'transform': 'title'}
        elif choice == '4':
            overrides[cname] = {'transform': 'upper'}
        elif choice == '5':
            pattern = input('    Enter regex pattern (e.g., ^EMP-\\d+$): ').strip()
            if pattern:
                rules.setdefault('regex_validations', {})[cname] = pattern
        elif choice == '6':
            overrides[cname] = {'ignore': True}

    rules['column_overrides'] = overrides

    save_name = input('\nSave this mapping as a reusable profile? Enter name (e.g., my_sales) or Enter to skip: ').strip()
    if save_name:
        if not save_name.endswith('.json'):
            save_name += '.json'
        save_path = os.path.join('profiles', save_name)
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(rules, f, indent=4)
        print(f'{GREEN}[SAVED] Reusable profile saved to: {save_path}{RESET}')

    return rules


# ==============================================================================
# 5. CORE CLEANING & VALIDATION ENGINE
# ==============================================================================
def format_row_preview(headers, row, max_cols=4):
    pairs = []
    for h, val in list(zip(headers, row))[:max_cols]:
        display_val = 'EMPTY' if is_empty_value(val) else str(val)
        pairs.append(f'{h}={display_val}')
    return ' | '.join(pairs)


def clean_dataset(headers, data, args, rules, schema_profile):
    initial_rows = len(data)
    cleaned_headers = standardize_headers(headers)
    col_index_map = {name: idx for idx, name in enumerate(cleaned_headers)}
    profile_map = {col['name']: col for col in schema_profile}

    # Resolve deduplication key indices
    dedup_indices = list(range(len(cleaned_headers)))
    if args.dedup_on:
        requested_keys = [k.strip().lower().replace(' ', '_') for k in args.dedup_on.split(',')]
        valid_keys = [col_index_map[k] for k in requested_keys if k in col_index_map]
        if valid_keys:
            dedup_indices = valid_keys
            print(f"{CYAN}[DEDUP] Deduplicating using key columns: {', '.join(requested_keys)} (keep={args.keep}){RESET}")

    print(f'\n{CYAN}{BOLD}=== STARTING DATA CLEANING AUDIT (Total Raw Rows: {initial_rows}) ==={RESET}\n')

    # Pre-pass if keeping 'last' duplicate
    last_occurrence_index = {}
    if args.drop_duplicates and args.keep == 'last':
        for idx, row in enumerate(data):
            norm = [c.strip() if isinstance(c, str) else c for c in row]
            if all(is_empty_value(c) for c in norm):
                continue
            key = tuple('' if (i >= len(norm) or is_empty_value(norm[i])) else str(norm[i]) for i in dedup_indices)
            last_occurrence_index[key] = idx

    cleaned_rows = []
    rejected_rows = []
    repaired_coords = set()
    seen = set()

    rejected_entries = []
    repaired_entries = []
    accepted_entries = []

    stats = {
        'accepted': 0,
        'repaired': 0,
        'removed_blank': 0,
        'removed_duplicate': 0,
        'removed_invalid': 0,
        'outliers_flagged': 0,
    }

    overrides = rules.get('column_overrides', {})
    regex_rules = rules.get('regex_validations', {})
    allowed_rules = rules.get('allowed_values', {})
    computed_rules = rules.get('computed_columns', [])

    for row_idx, raw_row in enumerate(data):
        excel_row_num = row_idx + 2
        current_row = list(raw_row[: len(cleaned_headers)])
        while len(current_row) < len(cleaned_headers):
            current_row.append(None)

        # 1. Completely Blank Row
        if all(is_empty_value(c) for c in current_row):
            stats['removed_blank'] += 1
            print(f'{RED}[Row {excel_row_num:03d}] REMOVED (BLANK ROW)   -> All cells are empty{RESET}')
            rejected_rows.append(current_row + ['Blank Row'])
            rejected_entries.append({
                'row': excel_row_num,
                'category': 'BLANK ROW',
                'preview': 'All cells are empty',
                'reason': 'Empty record',
            })
            continue

        # Trim whitespace
        current_row = [c.strip() if isinstance(c, str) else c for c in current_row]
        preview = format_row_preview(cleaned_headers, current_row)

        # 2. Duplicate Check (Full Row or Subset Keys)
        row_key = tuple('' if is_empty_value(current_row[i]) else str(current_row[i]) for i in dedup_indices)
        if args.drop_duplicates:
            is_dup = False
            if args.keep == 'first':
                if row_key in seen:
                    is_dup = True
                seen.add(row_key)
            else:  # keep == 'last'
                if last_occurrence_index.get(row_key) != row_idx:
                    is_dup = True

            if is_dup:
                stats['removed_duplicate'] += 1
                print(f'{RED}[Row {excel_row_num:03d}] REMOVED (DUPLICATE)   -> {preview}{RESET}')
                rejected_rows.append(current_row + [f'Duplicate Row (keep={args.keep})'])
                rejected_entries.append({
                    'row': excel_row_num,
                    'category': 'DUPLICATE',
                    'preview': preview,
                    'reason': f'Duplicate key match (keep={args.keep})',
                })
                continue

        row_repairs = []
        row_repaired_cols = set()

        # 3. Auto-Fix & Column Overrides Engine
        if args.auto_fix or overrides:
            for c_idx, col_name in enumerate(cleaned_headers):
                col_override = overrides.get(col_name, {})
                if col_override.get('ignore'):
                    continue

                val = current_row[c_idx]

                # Apply column-specific fill_value override
                if is_empty_value(val) and 'fill_value' in col_override:
                    current_row[c_idx] = col_override['fill_value']
                    row_repairs.append(f"{col_name}: EMPTY -> '{col_override['fill_value']}'")
                    row_repaired_cols.add(c_idx)
                    continue

                if is_empty_value(val):
                    continue

                # Apply explicit transform override
                if col_override.get('transform') == 'title':
                    new_v = ' '.join(str(val).split()).title()
                    if new_v != str(val):
                        current_row[c_idx] = new_v
                        row_repairs.append(f'{col_name}: {val} -> {new_v}')
                        row_repaired_cols.add(c_idx)
                    continue
                elif col_override.get('transform') == 'upper':
                    new_v = str(val).upper()
                    if new_v != str(val):
                        current_row[c_idx] = new_v
                        row_repairs.append(f'{col_name}: {val} -> {new_v}')
                        row_repaired_cols.add(c_idx)
                    continue

                if not args.auto_fix:
                    continue

                # Domain Keyword / Inferred Type Auto-Repairs
                inferred_type = profile_map.get(col_name, {}).get('type', 'Text')

                if inferred_type == 'Numeric' or any(
                    k in col_name for k in rules['currency_keywords'] + rules['quantity_keywords']
                ):
                    parsed = parse_number(val)
                    if parsed is not None and (parsed != val or type(parsed) is not type(val)):
                        current_row[c_idx] = parsed
                        row_repairs.append(f'{col_name}: {val} -> {parsed}')
                        row_repaired_cols.add(c_idx)

                elif inferred_type == 'Date' or any(k in col_name for k in rules['date_keywords']):
                    std_date, changed = standardize_date(val)
                    if changed:
                        current_row[c_idx] = std_date
                        row_repairs.append(f'{col_name}: {val} -> {std_date}')
                        row_repaired_cols.add(c_idx)

                elif inferred_type == 'Phone' or any(k in col_name for k in rules['phone_keywords']):
                    std_phone, changed = standardize_phone(val)
                    if changed:
                        current_row[c_idx] = std_phone
                        row_repairs.append(f'{col_name}: {val} -> {std_phone}')
                        row_repaired_cols.add(c_idx)

                elif inferred_type == 'Email' or any(k in col_name for k in rules['email_keywords']):
                    lower_email = str(val).lower()
                    if lower_email != str(val):
                        current_row[c_idx] = lower_email
                        row_repairs.append(f'{col_name}: {val} -> {lower_email}')
                        row_repaired_cols.add(c_idx)

                elif any(k in col_name for k in rules['name_keywords']) and 'product' not in col_name:
                    title_name = ' '.join(str(val).split()).title()
                    if title_name != str(val):
                        current_row[c_idx] = title_name
                        row_repairs.append(f'{col_name}: {val} -> {title_name}')
                        row_repaired_cols.add(c_idx)

                elif any(k in col_name for k in rules['city_keywords']):
                    city_clean = str(val).strip()
                    mapped = rules['city_aliases'].get(city_clean.lower(), city_clean.title())
                    if mapped != str(val):
                        current_row[c_idx] = mapped
                        row_repairs.append(f'{col_name}: {val} -> {mapped}')
                        row_repaired_cols.add(c_idx)

                elif any(k in col_name for k in rules['title_case_keywords']):
                    if str(val).lower() not in ['upi']:
                        title_stat = str(val).title()
                        if title_stat != str(val):
                            current_row[c_idx] = title_stat
                            row_repairs.append(f'{col_name}: {val} -> {title_stat}')
                            row_repaired_cols.add(c_idx)

            # 4. Dynamic Computed Formula Engine
            row_dict = dict(zip(cleaned_headers, current_row))
            for comp in computed_rules:
                target = comp.get('target')
                formula = comp.get('formula')
                only_empty = comp.get('only_if_empty', True)
                if target in col_index_map and formula:
                    t_idx = col_index_map[target]
                    if not only_empty or is_empty_value(current_row[t_idx]):
                        calc_val = safe_eval_formula(formula, row_dict)
                        if calc_val is not None and calc_val != current_row[t_idx]:
                            old_v = 'EMPTY' if is_empty_value(current_row[t_idx]) else current_row[t_idx]
                            current_row[t_idx] = calc_val
                            row_dict[target] = calc_val
                            row_repairs.append(f'{target}: {old_v} -> {calc_val} (Formula: {formula})')
                            row_repaired_cols.add(t_idx)

        # 5. Validation Engine (Missing, Placeholders, Regex, Allowed Enums, Statistical Outliers)
        issues = []
        for c_idx, (col_name, cell) in enumerate(zip(cleaned_headers, current_row)):
            col_override = overrides.get(col_name, {})
            if col_override.get('ignore'):
                continue

            if is_empty_value(cell):
                issues.append(f'{col_name}: Missing')
                continue

            val_str = str(cell).strip()

            # Invalid placeholder check
            if val_str.lower() in rules['invalid_placeholders']:
                issues.append(f"{col_name}: Invalid placeholder ('{val_str}')")

            # Email format check
            if any(k in col_name for k in rules['email_keywords']) and '@' not in val_str:
                issues.append(f"{col_name}: Invalid email ('{val_str}')")

            # Regex pattern check
            if col_name in regex_rules:
                if not re.match(regex_rules[col_name], val_str):
                    issues.append(f"{col_name}: Failed regex {regex_rules[col_name]} ('{val_str}')")

            # Allowed enum values check
            if col_name in allowed_rules:
                allowed = allowed_rules[col_name]
                if val_str not in allowed:
                    issues.append(f"{col_name}: '{val_str}' not in {allowed}")

            # Numeric & Statistical Outlier check
            if any(k in col_name for k in rules['currency_keywords'] + rules['quantity_keywords']):
                num = parse_number(cell)
                if num is None or num < 0:
                    issues.append(f"{col_name}: Invalid number ('{val_str}')")
                elif args.flag_outliers:
                    cstats = profile_map.get(col_name, {}).get('stats', {})
                    if cstats.get('stdev', 0) > 0:
                        z_score = abs(num - cstats['mean']) / cstats['stdev']
                        if z_score > 3.0:
                            stats['outliers_flagged'] += 1
                            issues.append(f"{col_name}: Statistical Outlier ({num}, z={z_score:.1f}σ)")

        updated_preview = format_row_preview(cleaned_headers, current_row)

        if issues:
            issue_str = ', '.join(issues)
            # Drop if --drop-na is set OR if any failing column was marked required in column_overrides
            has_required_fail = any(
                overrides.get(iss.split(':')[0], {}).get('required') for iss in issues
            )
            if args.drop_na or has_required_fail:
                stats['removed_invalid'] += 1
                print(f'{RED}[Row {excel_row_num:03d}] REMOVED (BAD/MISSING) -> {updated_preview}{RESET}')
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
                        row_repairs.append(f"{cleaned_headers[c_idx]}: EMPTY -> '{args.fill_na}'")
                dest_row_idx = len(cleaned_rows)
                for c_idx in row_repaired_cols:
                    repaired_coords.add((dest_row_idx, c_idx))
                cleaned_rows.append(current_row)
                stats['repaired'] += 1
                print(f"{YELLOW}[Row {excel_row_num:03d}] FIXED (FILLED '{args.fill_na}') -> {updated_preview}{RESET}")
                print(f'{YELLOW}          └── Issues Handled: {issue_str}{RESET}')
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
                print(f'{RED}[Row {excel_row_num:03d}] WARNING (MISSING/WRONG) -> {updated_preview}{RESET}')
                print(f'{RED}          └── Remaining Issues: {issue_str}{RESET}')
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
            print(f'{YELLOW}[Row {excel_row_num:03d}] INFO: REPAIRED & ACCEPTED -> {updated_preview}{RESET}')
            print(f"{YELLOW}          └── Auto-Fixed: {', '.join(row_repairs)}{RESET}")
            repaired_entries.append({
                'row': excel_row_num,
                'preview': updated_preview,
                'fixes': row_repairs,
            })
        else:
            stats['accepted'] += 1
            print(f'{GREEN}[Row {excel_row_num:03d}] INFO: ACCEPTED            -> {updated_preview}{RESET}')
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
    if args.flag_outliers:
        print(f"{MAGENTA}⚠ Outliers Detected (>3σ) : {stats['outliers_flagged']}{RESET}")
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


# ==============================================================================
# 6. AUDIT LOG GENERATOR
# ==============================================================================
def write_professional_log(
    log_path,
    input_path,
    sheet_name,
    output_path,
    quarantine_path,
    args,
    rules,
    stats,
    initial_rows,
    final_clean_count,
    rejected_count,
    rejected_entries,
    repaired_entries,
    accepted_entries,
    duration_sec,
):
    run_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sep_heavy = '=' * 94
    sep_light = '-' * 94

    def pct(count):
        return f'{(count / initial_rows * 100):.1f}%' if initial_rows else '0.0%'

    lines = [
        sep_heavy,
        'ENTERPRISE DATA CLEANING PIPELINE - EXECUTION AUDIT LOG',
        sep_heavy,
        f'Execution Timestamp   : {run_timestamp}',
        f'Execution Duration    : {duration_sec:.2f} seconds',
        f'Active Domain Profile : {rules.get("domain", "custom").upper()}',
        f'Source Input File     : {os.path.abspath(input_path)} (Sheet: {sheet_name or "Default"})',
        f'Cleaned Output File   : {os.path.abspath(output_path)}',
        f'Quarantine Output File: {os.path.abspath(quarantine_path) if rejected_count else "N/A"}',
        f'Pipeline Flags        : auto_fix={args.auto_fix} | drop_duplicates={args.drop_duplicates} (keep={args.keep}) | drop_na={args.drop_na} | flag_outliers={args.flag_outliers}',
        sep_heavy,
        '',
        '1. EXECUTIVE SUMMARY',
        sep_light,
        f"{'Metric':<34} | {'Row Count':<12} | {'Share (%)':<10}",
        sep_light,
        f"{'Total Raw Rows Processed':<34} | {initial_rows:<12} | {'100.0%':<10}",
        f"{'Clean Accepted (Unmodified)':<34} | {stats['accepted']:<12} | {pct(stats['accepted']):<10}",
        f"{'Auto-Repaired & Accepted':<34} | {stats['repaired']:<12} | {pct(stats['repaired']):<10}",
        f"{'Removed - Blank Rows':<34} | {stats['removed_blank']:<12} | {pct(stats['removed_blank']):<10}",
        f"{'Removed - Duplicate Rows':<34} | {stats['removed_duplicate']:<12} | {pct(stats['removed_duplicate']):<10}",
        f"{'Removed - Invalid/Missing Data':<34} | {stats['removed_invalid']:<12} | {pct(stats['removed_invalid']):<10}",
        sep_light,
        f"{'FINAL SAVED CLEAN ROWS':<34} | {final_clean_count:<12} | {pct(final_clean_count):<10}",
        f"{'TOTAL QUARANTINED ROWS':<34} | {rejected_count:<12} | {pct(rejected_count):<10}",
        sep_heavy,
        '',
        f'2. REJECTED & QUARANTINED RECORDS ({len(rejected_entries)} rows)',
        sep_light,
    ]

    if rejected_entries:
        for entry in rejected_entries:
            lines.append(f"[Row {entry['row']:03d}] [{entry['category']}] {entry['preview']}")
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


# ==============================================================================
# 7. INTERACTIVE WIZARD & ORCHESTRATION
# ==============================================================================
def run_interactive_wizard(args):
    print(f'\n{CYAN}{BOLD}=== UNIVERSAL DATA CLEANING WIZARD ==={RESET}')
    inp = input(f'1. Input file path [{args.input}]: ').strip()
    if inp:
        args.input = inp

    print('2. Available Domain Presets: [general, ecommerce, hr, finance, healthcare]')
    dom = input(f'   Select domain preset [{args.domain}]: ').strip().lower()
    if dom in DOMAIN_PRESETS:
        args.domain = dom

    fmt = input(f'3. Output format (xlsx / csv / json) [{args.format}]: ').strip().lower()
    if fmt in ('xlsx', 'csv', 'json'):
        args.format = fmt

    base_out = os.path.splitext(args.output)[0]
    args.output = f'{base_out}.{args.format}'
    base_quar = os.path.splitext(args.quarantine)[0]
    args.quarantine = f'{base_quar}.{args.format}'

    args.drop_duplicates = input('4. Remove duplicate rows? (y/n) [y]: ').strip().lower() != 'n'
    if args.drop_duplicates:
        dedup_cols = input('   Deduplicate on specific key columns? (comma-separated or Enter for full row): ').strip()
        if dedup_cols:
            args.dedup_on = dedup_cols
        keep_choice = input('   Keep [first] or [last] duplicate? [first]: ').strip().lower()
        if keep_choice in ('first', 'last'):
            args.keep = keep_choice

    args.auto_fix = input('5. Auto-repair formats, dates, phones, cities & formulas? (y/n) [y]: ').strip().lower() != 'n'
    args.flag_outliers = input('6. Flag statistical numeric outliers (>3σ)? (y/n) [n]: ').strip().lower() == 'y'
    args.map_columns = input('7. Launch interactive column-by-column rule mapper? (y/n) [n]: ').strip().lower() == 'y'

    na_choice = input('8. Missing values policy: [1] Keep & Flag  [2] Drop & Quarantine  [3] Fill with value (default 2): ').strip()
    if na_choice == '1':
        args.drop_na = False
    elif na_choice == '3':
        args.drop_na = False
        args.fill_na = input('   Enter replacement text [Unknown]: ').strip() or 'Unknown'
    else:
        args.drop_na = True

    args.dry_run = input('9. Run in Dry-Run preview mode (no files written)? (y/n) [n]: ').strip().lower() == 'y'


def process_sheet(input_path, sheet_name, output_path, quarantine_path, args, rules):
    start_time = time.time()
    headers, data = load_data(input_path, sheet_name)
    _, schema_profile = profile_schema(headers, data)
    print_schema_table(schema_profile)

    if args.profile_only:
        return

    if args.map_columns:
        rules = run_column_mapper(schema_profile, rules)

    (
        cleaned_headers,
        cleaned_data,
        rejected_data,
        repaired_coords,
        stats,
        rejected_entries,
        repaired_entries,
        accepted_entries,
    ) = clean_dataset(headers, data, args, rules, schema_profile)

    if args.dry_run:
        print(f'{MAGENTA}{BOLD}[DRY-RUN] Preview complete. No output or quarantine files were written.{RESET}')
        return

    save_output(
        cleaned_headers,
        cleaned_data,
        output_path,
        sheet_name='Cleaned_Data',
        repaired_coords=repaired_coords,
    )
    print(f'{GREEN}[SAVED] Cleaned dataset written to: {output_path}{RESET}')

    if rejected_data and quarantine_path:
        rejected_headers = cleaned_headers + ['rejection_reason']
        save_output(
            rejected_headers,
            rejected_data,
            quarantine_path,
            sheet_name='Rejected_Data',
            is_rejected=True,
        )
        print(f'{RED}[QUARANTINE] {len(rejected_data)} removed rows saved to: {quarantine_path}{RESET}')

    duration_sec = time.time() - start_time

    if args.log_file:
        log_path = args.log_file
    else:
        base_dir = os.path.dirname(os.path.abspath(input_path))
        logs_dir = os.path.join(base_dir, 'Logs')
        input_stem = os.path.splitext(os.path.basename(input_path))[0]
        sheet_tag = f'_{sheet_name}' if sheet_name else ''
        ts = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        log_path = os.path.join(logs_dir, f'audit_{input_stem}{sheet_tag}_{ts}.log')

    write_professional_log(
        log_path=log_path,
        input_path=input_path,
        sheet_name=sheet_name,
        output_path=output_path,
        quarantine_path=quarantine_path,
        args=args,
        rules=rules,
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


def process_file(input_path, output_path, quarantine_path, args, rules):
    ext = os.path.splitext(input_path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        sheets = get_excel_sheets(input_path)
        if len(sheets) > 1:
            if args.all_sheets:
                for sh in sheets:
                    out_sh = f'{os.path.splitext(output_path)[0]}_{sh}.{args.format}'
                    rej_sh = f'{os.path.splitext(quarantine_path)[0]}_{sh}.{args.format}'
                    print(f'\n{CYAN}{BOLD}>>> Cleaning Sheet: {sh}{RESET}')
                    process_sheet(input_path, sh, out_sh, rej_sh, args, rules)
                return
            elif args.sheet:
                process_sheet(input_path, args.sheet, output_path, quarantine_path, args, rules)
                return
            else:
                print(f'\n{CYAN}Workbook contains multiple sheets: {sheets}{RESET}')
                chosen = input(f'Enter sheet name to clean [{sheets[0]}]: ').strip() or sheets[0]
                process_sheet(input_path, chosen, output_path, quarantine_path, args, rules)
                return

    process_sheet(input_path, args.sheet, output_path, quarantine_path, args, rules)


def main():
    parser = argparse.ArgumentParser(
        description='Universal Multi-Domain CLI Data Cleaning & Profiling Tool (Pure Python).'
    )

    parser.add_argument('input', nargs='?', default='raw_data.xlsx', help='Input file (CSV, Excel, JSON)')
    parser.add_argument('-o', '--output', default='cleaned_data.xlsx', help='Cleaned output file path')
    parser.add_argument('-q', '--quarantine', default='rejected_data.xlsx', help='Quarantine output file path')
    parser.add_argument('--format', choices=['xlsx', 'csv', 'json'], default='xlsx', help='Output format')
    parser.add_argument('--domain', choices=list(DOMAIN_PRESETS.keys()), default='general', help='Industry preset profile')
    parser.add_argument('--profile', type=str, default=None, help='Path to custom JSON profile')
    parser.add_argument('--profile-only', action='store_true', help='Inspect schema & data health without cleaning')
    parser.add_argument('--map-columns', action='store_true', help='Launch interactive column-by-column rule mapper')
    parser.add_argument('--drop-duplicates', action='store_true', help='Remove duplicate rows')
    parser.add_argument('--dedup-on', type=str, default=None, help='Comma-separated key columns for deduplication')
    parser.add_argument('--keep', choices=['first', 'last'], default='first', help='Which duplicate occurrence to keep')
    parser.add_argument('--auto-fix', action='store_true', help='Auto-repair types, dates, phones, cities & formulas')
    parser.add_argument('--drop-na', action='store_true', help='Drop rows with invalid/missing values into quarantine')
    parser.add_argument('--fill-na', type=str, default=None, help='Fill missing cells with a default value')
    parser.add_argument('--flag-outliers', action='store_true', help='Flag numeric values > 3 standard deviations')
    parser.add_argument('--sheet', type=str, default=None, help='Specific Excel worksheet name to clean')
    parser.add_argument('--all-sheets', action='store_true', help='Clean every worksheet in an Excel workbook')
    parser.add_argument('--folder', type=str, default=None, help='Batch clean all files in a directory')
    parser.add_argument('--log-file', type=str, default=None, help='Custom path for audit log')
    parser.add_argument('--dry-run', action='store_true', help='Preview audit without saving output files')
    parser.add_argument('-i', '--interactive', action='store_true', help='Launch interactive setup wizard')

    args = parser.parse_args()

    if args.interactive:
        run_interactive_wizard(args)

    # Sync output/quarantine extensions with --format if user didn't specify custom extensions
    if not args.output.endswith(f'.{args.format}'):
        args.output = f'{os.path.splitext(args.output)[0]}.{args.format}'
    if not args.quarantine.endswith(f'.{args.format}'):
        args.quarantine = f'{os.path.splitext(args.quarantine)[0]}.{args.format}'

    rules = load_rules(args)

    if args.folder:
        if not os.path.isdir(args.folder):
            print(f"{RED}Error: Folder '{args.folder}' does not exist.{RESET}")
            sys.exit(1)
        out_dir = os.path.join(args.folder, 'cleaned_output')
        os.makedirs(out_dir, exist_ok=True)
        files = [f for f in os.listdir(args.folder) if f.lower().endswith(('.csv', '.xlsx', '.xls', '.json'))]
        if not files:
            print(f"{YELLOW}No supported files found in '{args.folder}'.{RESET}")
            return
        for fname in files:
            stem = os.path.splitext(fname)[0]
            in_p = os.path.join(args.folder, fname)
            out_p = os.path.join(out_dir, f'cleaned_{stem}.{args.format}')
            rej_p = os.path.join(out_dir, f'rejected_{stem}.{args.format}')
            print(f'\n{CYAN}{BOLD}>>> Processing Batch File: {fname}{RESET}')
            process_file(in_p, out_p, rej_p, args, rules)
    else:
        process_file(args.input, args.output, args.quarantine, args, rules)


if __name__ == '__main__':
    main()