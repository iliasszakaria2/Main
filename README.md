# Main

This repository contains a Python utility script for updating the **PDP** Excel workbook
using data from the **Prévision EMEA** workbook.

## Usage

```bash
pip install pandas openpyxl
python process_prevision.py PDP.xlsx Prevision_EMEA.xlsx
```

The script will:

1. Filter rows in `Prévision EMEA` sheet **Feuil 1** where:
   - `Fiability rate` is between 60% and 90% (inclusive).
   - `EnerOne B-Cab` is not empty.
2. Sum the values of `EnerOne B-Cab` for each `Livraison` date.
3. Insert the sums into the `B-CAB` sheet of `PDP.xlsx` under
   `Customer Opportunities` of `B-CAB L E 0.5C`, matching the dates listed on
   row 78 of the sheet.

The provided workbooks are modified in place.
