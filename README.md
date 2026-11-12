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
