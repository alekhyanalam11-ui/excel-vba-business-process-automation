
# Excel VBA Business Process Automation — Rice Trading Operations

## Project overview

This repository documents an Excel VBA automation project used in a rice trading/canvassing business workflow. It contains three standard VBA modules and worksheet event-procedure source files transcribed from the VBA reference document supplied for this project.

The code covers commission-rate checking, Prakasam-region report extraction, pending-delivery reporting, transaction-sheet event handling, party/miller statements, and area-wise summary handlers. The repository is source code and documentation; it is not a ready-to-run Excel workbook.

## Main capabilities represented in the source

- **Commission-rate checking:** compares transaction-sheet commission rates with the PARTY master list and highlights mismatches.
- **Prakasam reports:** extracts Prakasam-area pending-delivery and pending-commission rows from NONLOCAL into PRAKASAM.
- **Pending-delivery reports:** refreshes pending-delivery output from LOCAL and NONLOCAL and provides party-group extraction procedures.
- **Transaction and account workflows:** worksheet events perform lookups/calculations and generate party-, miller-, and results-sheet reports.
- **Area summaries:** worksheet event code is included separately for each of the 14 area worksheets.

These are descriptions of code paths present in the supplied source. They are not a claim that every workflow has been runtime-tested in a clean Excel environment.

## Repository structure

```text
excel-vba-business-process-automation/
├── README.md
├── .gitignore
├── docs/
│   ├── WORKBOOK_STRUCTURE.md
│   ├── VBA_INVENTORY.md
│   ├── WORKFLOW_OVERVIEW.md
│   └── TESTING_CHECKLIST.md
└── src/
    ├── modules/
    │   ├── modCommissionValidation.bas
    │   ├── modPrakasamReports.bas
    │   └── modPendingDelivery.bas
    └── worksheets/
        ├── Sheet_LOCAL.cls
        ├── Sheet_NONLOCAL.cls
        ├── Sheet_PARTY.cls
        ├── Sheet_MILLER.cls
        ├── Sheet_Results_LOCAL.cls
        ├── Sheet_Results_NONLOCAL.cls
        ├── Sheet_MILLER_LOCAL.cls
        ├── Sheet_Results_MILLER_LOCAL.cls
        ├── Sheet_PENDING_DELIVERY.cls
        ├── Sheet_PRAKASAM.cls
        └── 14 area worksheet code files
```

## Using the source files

1. Make a backup copy of the workbook and open it in desktop Microsoft Excel.
2. Press `Alt + F11` to open the Visual Basic Editor.
3. For standard modules, use **File → Import File...** and import the `.bas` files from `src/modules/` into the VBA project. Check for existing modules with the same names before importing to avoid duplicate procedures.
4. Worksheet event code must be placed in the matching worksheet's code window (under **Microsoft Excel Objects** in Project Explorer). Open the corresponding `.cls` source file as text, copy its procedure code, and paste it into the matching worksheet code window. Do not import a worksheet event file as a normal class module; that would not attach the events to the worksheet.
5. Save a separate test copy and run the relevant checks in `docs/TESTING_CHECKLIST.md`.

The 14 area-sheet code files are individual worksheet handlers. They are not a shared helper module; retain each handler in its corresponding worksheet module unless you intentionally undertake and test a refactor.

## Source and verification notes

- The worksheet and standard-module source files were transcribed from the supplied `VBA(2).docx` reference. They are not a newly optimized or independently rewritten version of the macros.
- The accompanying demo workbook contains 24 worksheet names; the sheet inventory is in `docs/WORKBOOK_STRUCTURE.md`.
- The workbook itself is intentionally not included in this repository package.
- The demo workbook was described as using demonstration data. Before publishing, independently confirm that names, balances, dates, screenshots, and any other identifying or commercially sensitive information are safe to disclose.
- No claim is made here that the code has passed a formal test suite or been compiled in every target Excel version.
