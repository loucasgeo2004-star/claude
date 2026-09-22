# Excel Formula Re-Pointer for Trident Unity LP

Automatically re-points BS and IS formulas from T sheet to TB2026 sheet.

## What It Does

This script takes an Excel file and automatically:
1. ✅ Scans TB2026 to build a mapping of T cells → TB2026 cells
2. ✅ Updates all BS (Balance Sheet) column I formulas
3. ✅ Updates all IS (Income Statement) column I formulas
4. ✅ Saves a new file with `_REPOINTED` suffix
5. ✅ Displays a summary of all changes

**No manual work needed** — just run it once per file!

---

## Installation

### Requirements
- Python 3.7 or higher
- openpyxl library

### Setup (One time)

```bash
# Install openpyxl if not already installed
pip install openpyxl

# Make the script executable (macOS/Linux)
chmod +x excel_formula_repointer.py
```

---

## Usage

### Option 1: Command Line (Recommended)

```bash
python3 excel_formula_repointer.py /path/to/your/file.xlsx
```

**Example:**
```bash
python3 excel_formula_repointer.py "Accounting_2026.xlsx"
```

### Option 2: Double-Click (macOS/Linux)

Add this to the top of the script (already included):
```python
#!/usr/bin/env python3
```

Then you can double-click the script and it will prompt for a file.

### Option 3: Drag & Drop

```bash
# Save this as a helper script or alias
python3 excel_formula_repointer.py "$1"
```

---

## Output

The script creates a new file with `_REPOINTED` suffix:

**Input:** `Accounting_2026.xlsx`  
**Output:** `Accounting_2026_REPOINTED.xlsx`

The original file is **never modified**.

---

## Example Run

```
================================================================================
TRIDENT UNITY LP - FORMULA RE-POINTER
================================================================================

Input file: Accounting_2026.xlsx
Timestamp: 2026-09-22 11:29:12

📂 Loading workbook...
✓ All required sheets found (TB2026, BS, IS, T)

🔍 Building T → TB2026 mapping...
✓ Found 41 T cell → TB2026 mappings

📊 Updating BS (Balance Sheet)...
✓ Updated 8 BS formulas

📈 Updating IS (Income Statement)...
✓ Updated 17 IS formulas

💾 Saving updated file...
✓ Saved to: Accounting_2026_REPOINTED.xlsx

================================================================================
✅ DONE!
```

---

## What Gets Updated

### Balance Sheet (typically 8 cells)
- Individual account values from T → TB2026
- Aggregate calculations stay at T (no equivalent in TB2026)

### Income Statement (typically 17 cells)
- Individual account values from T → TB2026
- Aggregate calculations stay at T (no equivalent in TB2026)

**Total: ~25 cells typically updated**

---

## Verification Steps

After running the script:

1. **Open the output file** in Excel or LibreOffice
2. **Recalculate** (Ctrl+Shift+F9 or Tools → Recalculate)
3. **Compare values** with the original file
   - All BS/IS column I values should be identical
   - Balance Sheet should still balance (Assets = Liabilities + Equity)
4. **Done!** ✅

---

## How It Works (Technical)

1. **Loads TB2026 sheet** and scans all D and E columns for formulas
2. **Extracts T cell references** (e.g., `=T!B12` → saves `T!B12`)
3. **Creates mapping** (T!B12 → TB2026!E9, etc.)
4. **Scans BS and IS column I** for any formulas with T references
5. **Replaces matching T references** with TB2026 equivalents
6. **Leaves unmapped cells** (aggregate net calculations) pointing to T
7. **Saves new file** with all updates

---

## File Requirements

Your Excel file **must have these sheets:**
- ✅ TB2026 (Trial Balance 2026)
- ✅ BS (Balance Sheet)
- ✅ IS (Income Statement)
- ✅ T (Transactions/Calculations)
- ACC ENTRIES 2026 (not required, but recommended)

---

## Troubleshooting

### "File not found"
Check the file path is correct:
```bash
python3 excel_formula_repointer.py "./Accounting_2026.xlsx"
```

### "Required sheet not found"
Your file doesn't have all required sheets (TB2026, BS, IS, T). Verify the sheet names.

### "No updates applied"
Either:
- All formulas are already correct
- Or the file structure is different than expected

### Python not found
Install Python 3 from python.org or use `python` instead of `python3`

---

## Support

If something goes wrong, check:
1. Python version: `python3 --version` (should be 3.7+)
2. openpyxl installed: `pip install openpyxl`
3. File exists and is a valid Excel file
4. All required sheets are present (TB2026, BS, IS, T)

---

## License

Made for Trident Unity LP accounting automation.
