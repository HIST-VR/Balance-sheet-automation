# Excel Accounting Tools

Small Python tools for processing 1C "turnover balance sheet" Excel exports for print-shop orders. They work with both `.xls` and `.xlsx`, never modify the original file, and mark anything they can't calculate.

## Tools

**1. `Excel_Accounting.py`** – calculates quantities and categories
- Reads the run size from the item name (`Business card, 1000 pcs`) and multiplies it by the turnover. The result goes into column D.
- Items without a "шт" (pcs) in the name count as 1 piece each.
- Assigns a category number to column C using `reference.txt` (singular and plural forms are matched). Categories 12–14 are highlighted yellow.
- Writes the grand total in the "Разом" row.
- Ambiguous rows are marked `########` and listed in a summary dialog.

**2. `Excel_Summary.py`** – builds the per-category summary block
- Takes the output of tool 1 and appends a summary block below the last row.
- Per category: total quantity, total amount and weight formulas, plus the rest of the report template (colors included).

## Requirements
- Python 3.9+
- `pip install openpyxl xlrd xlutils xlwt`

## Usage
Drag an Excel file onto the script, or run it and pick a file in the dialog. Run tool 1 first, then tool 2 on its output.

To get standalone `.exe` files that don't need Python, run `Build_exe.bat` once on Windows.

## Category reference
`reference.txt` maps category numbers to product words. Put it next to the program and edit it to add new words:

```
3 - візитка, квиток, табличка
10 - наклейка, стікер
```

If `reference.txt` is missing, a built-in copy is used.

## Notes
- Both tools stop with a clear message
- Columns are found by their headers, not by fixed positions.
