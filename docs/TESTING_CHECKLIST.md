# Testing Checklist

> This is a proposed test checklist. It is not evidence that the tests have been executed or passed. Record the test date, Excel version, macro settings, input case, expected result, actual result, and pass/fail status when testing.

Run tests only on a backup copy of the workbook.

## Pre-test checks
- [ ] Confirm all 24 expected worksheets exist.
- [ ] Confirm PARTY and MILLER master data is populated as required.
- [ ] Confirm macros are enabled in the test copy.
- [ ] Confirm the workbook has been backed up and can be restored.

## Commission-rate validation
- [ ] Test a matching commission rate.
- [ ] Test a deliberately incorrect commission rate.
- [ ] Test a party name that is blank or absent from the PARTY master list.
- [ ] Confirm highlighting is applied only to the intended rate cells.

## Prakasam reports
- [ ] Test a Prakasam party with a blank delivery date.
- [ ] Test a Prakasam party with a valid delivery date and positive P DEBIT.
- [ ] Test zero, negative, blank, and nonnumeric debit values.
- [ ] Confirm output headers, sort order, and report boundaries.

## Pending-delivery reports
- [ ] Test qualifying LOCAL rows.
- [ ] Test qualifying NONLOCAL rows.
- [ ] Test each party-group extraction procedure.
- [ ] Test excluded party prefixes and similarly named nonmatching parties.
- [ ] Test no qualifying rows and repeated runs.
- [ ] Confirm output clearing does not affect unrelated cells.

## Worksheet event handlers
- [ ] Test party and miller lookups using valid and invalid inputs.
- [ ] Test relevant calculation columns with normal, blank, zero, and negative values.
- [ ] Test date-range boundaries and ranges with no matching rows.
- [ ] Test each party/miller and results-sheet report trigger.
- [ ] Test each of the 14 area-sheet summary handlers.
- [ ] Test single-cell edits and multi-cell paste operations.

## Error and recovery checks
- [ ] Trigger a controlled error in a disposable test copy and observe the error handling.
- [ ] Verify Excel events and screen updating are left in a usable state after success and failure.
- [ ] Confirm no source transaction rows are unexpectedly deleted or overwritten.

## Release gate

Do not describe a test as passed until the actual output has been compared with the source rows and the result has been recorded. Do not publish screenshots or data extracts containing information that should remain private.
