# Workbook Structure

## Worksheet inventory

The supplied demo workbook (`Github Demo(1).xlsb`) was inspected for its worksheet names. It contains 24 worksheets:

- `LOCAL`
- `NONLOCAL`
- `PARTY`
- `MILLER`
- `Results_LOCAL`
- `Results_NONLOCAL`
- `PENDING_DELIVERY`
- `PRAKASAM`
- `MILLER_LOCAL`
- `Results_MILLER_LOCAL`
- `ELR ROAD`
- `SINGHNAGAR`
- `SN PURAM`
- `KOTHAPETA`
- `PATAMATA`
- `ELR`
- `NZV`
- `PNR`
- `TNL`
- `BHVM`
- `MGR`
- `BANDAR`
- `GUNTUR`
- `TIRUVURU`

## Worksheet code coverage in this repository

The source reference contains worksheet event code for the following worksheets:

- `LOCAL`
- `NONLOCAL`
- `MILLER`
- `PARTY`
- `Results_LOCAL`
- `Results_NONLOCAL`
- `MILLER_LOCAL`
- `Results_MILLER_LOCAL`
- `PENDING_DELIVERY`
- `PRAKASAM`
- `KOTHAPETA`
- `PATAMATA`
- `ELR`
- `NZV`
- `PNR`
- `TNL`
- `BHVM`
- `MGR`
- `BANDAR`
- `GUNTUR`
- `TIRUVURU`
- `ELR ROAD`
- `SINGHNAGAR`
- `SN PURAM`


The 14 area sheets are represented by individual event-code files. The source document uses the labels `KOTTHAPETA` and `SNPURAM`; the filenames here use the actual workbook sheet names `KOTHAPETA` and `SN PURAM` respectively.

## Scope and limitations

- This inventory records worksheet names, not every column, formula, named range, table, or control in the workbook.
- The code files are source snapshots from the supplied VBA reference document. They should be checked against the workbook's current VBA project before deployment if the workbook has changed since that document was prepared.
- The `.xlsb` demo workbook is not included in this source package.
