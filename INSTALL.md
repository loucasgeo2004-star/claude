# Installation Guide - Excel Formula Re-Pointer

Quick setup for your laptop in **2 minutes**.

---

## Installation (macOS & Linux)

### Step 1: Download the Files

Clone or download these 3 files to a folder on your laptop:
- `excel_formula_repointer.py`
- `setup_formula_repointer.sh`
- `README_FORMULA_REPOINTER.md`

### Step 2: Run the Setup Script

Open Terminal and navigate to the folder:

```bash
cd /path/to/folder/with/scripts
```

Then run:

```bash
bash setup_formula_repointer.sh
```

That's it! The script will:
- ✅ Check Python (install if needed)
- ✅ Install openpyxl library
- ✅ Copy the tool to system location
- ✅ Add it to your PATH

### Step 3: Restart Terminal (if needed)

Close and reopen Terminal, or run:

```bash
source ~/.zshrc
# OR for bash:
source ~/.bashrc
```

---

## Usage (After Installation)

Now you can run from anywhere:

```bash
excel_formula_repointer your_file.xlsx
```

**Examples:**

```bash
# File in same folder
excel_formula_repointer Trident_2026.xlsx

# File in Documents
excel_formula_repointer ~/Documents/Trident_2026.xlsx

# Full path
excel_formula_repointer /Users/yourname/Desktop/Accounting.xlsx
```

---

## What Happens

1. **Input:** `Trident_2026.xlsx`
2. **Process:** Updates all formulas automatically
3. **Output:** `Trident_2026_REPOINTED.xlsx` (new file created)

The original file is never modified.

---

## Verification

After running:

1. Open the `_REPOINTED.xlsx` file
2. Press `Ctrl+Shift+F9` (or `Cmd+Shift+F9` on Mac) to recalculate
3. Compare numbers with original file
4. Done! ✅

---

## Troubleshooting

### "Command not found: excel_formula_repointer"

**Solution:** The script wasn't added to PATH. Run setup again:
```bash
bash setup_formula_repointer.sh
```

Then restart Terminal.

### "Python not found"

**Solution:** Install Python from https://www.python.org

Then try setup again.

### "Permission denied"

**Solution:** Make setup script executable:
```bash
chmod +x setup_formula_repointer.sh
bash setup_formula_repointer.sh
```

---

## For Windows Users

Windows doesn't have a native bash shell, but you can:

**Option 1: Use Python directly**
```bash
python3 excel_formula_repointer.py C:\path\to\file.xlsx
```

**Option 2: Install WSL (Windows Subsystem for Linux)**
Then follow the macOS/Linux instructions above.

**Option 3: Use Command Prompt**
```bash
python3 excel_formula_repointer.py "C:\Users\YourName\Documents\file.xlsx"
```

---

## After Installation

- Run `excel_formula_repointer --help` for usage info
- Put your Excel files anywhere on your computer
- Run the command from any Terminal location
- Repeat as needed for new files

---

## Support

If something doesn't work:

1. Check Python: `python3 --version` (should be 3.7+)
2. Check openpyxl: `python3 -m pip list | grep openpyxl`
3. Run setup again: `bash setup_formula_repointer.sh`
4. Restart Terminal

---

**That's it! You're all set.** 🎉
