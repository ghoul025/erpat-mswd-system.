<img width="1322" height="588" alt="Screenshot 2026-07-16 010606" src="https://github.com/user-attachments/assets/416acf85-f624-4ab7-8ae7-4886e7b3a62a" />
<img width="1355" height="709" alt="Screenshot 2026-07-16 000125" src="https://github.com/user-attachments/assets/010b05e9-0d0f-4cf0-853f-adb9277c0804" />
<img width="727" height="571" alt="Screenshot 2026-07-16 000023" src="https://github.com/user-attachments/assets/9f52c631-60a5-401e-a141-48d00b0f3ab0" />
<img width="1340" height="630" alt="Screenshot 2026-07-15 235942" src="https://github.com/user-attachments/assets/33bac098-0b72-4630-8986-214435330d70" />
<img width="778" height="552" alt="Screenshot 2026-07-15 235541" src="https://github.com/user-attachments/assets/37edea3e-e6b6-4c57-9d79-6e28122e2d6d" />
# MSWD Registry System

An Excel VBA-based registry and ID management system built for the **Municipal Social Welfare and Development (MSWD)** office to replace manual record-keeping and ID processing.

## Problem It Solves

The MSWD office previously relied on manual, paper-based processes for managing beneficiary records and producing printed IDs — slow, error-prone, and hard to keep consistent. This system centralizes that into a single Excel-based application with a controlled, form-driven interface, so records can only be created, edited, or deleted through the proper UI rather than the raw spreadsheet.

## Features

- **CRUD UserForm interface** — all record creation, editing, and deletion goes through a dedicated form; the underlying database is locked from direct/manual editing to preserve data integrity
- **Live keyword filtering** — search and filter records in real time as you type, directly from the CRUD form
- **Automatic calendar generation** — built-in calendar view generated from the system
- **Auto ID generation with duplicate checking** — automatically assigns unique IDs to new records and validates against duplicates before saving
- **Front-and-back ID card generation** — select any record from the table and the system auto-generates a print-ready, complete front-and-back ID layout
- **Dashboard analytics** — summary view of registry data for quick monitoring and reporting

## Tech Stack

- **Platform:** Microsoft Excel
- **Automation:** VBA (UserForms, macros)

## Impact

Used by the Municipal Social Welfare and Development office to replace manual registry and ID-processing work with a controlled, automated system — reducing data entry errors and speeding up ID issuance.

## Setup / How to Use

```
1. Download/clone the workbook file from this repository.
2. Open the .xlsm file in Microsoft Excel.
3. Enable macros when prompted:
   File > Options > Trust Center > Trust Center Settings > Macro Settings > Enable all macros
4. Use the CRUD UserForm (accessible via the main menu/button on open) to add,
   search, edit, or delete records — do not edit the database sheet directly.
5. Select a record from the table and use the ID generation feature to produce
   a print-ready front-and-back ID card.
```

*(Adjust the exact filename and menu access steps above to match your actual workbook.)*

## Author

**Errol L. Alday**
[github.com/ghoul025](https://github.com/ghoul025)
