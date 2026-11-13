# Workflow Overview

This document summarizes the major code areas visible in the supplied VBA source. It is a source-level overview, not proof of successful runtime behavior.

## 1. Commission-rate check

`CheckAllCommissionRates` reads party names and commission rates from the PARTY sheet, then checks the corresponding rate cells in LOCAL and NONLOCAL. Mismatches are highlighted for review.

## 2. Prakasam reports

The Prakasam procedures identify party names associated with the configured Prakasam-area prefixes and extract qualifying NONLOCAL rows into PRAKASAM report sections. One report targets rows with a blank delivery date; the other filters on delivery date and positive P DEBIT.

## 3. Pending-delivery reports

The pending-delivery module refreshes output from LOCAL and NONLOCAL based on blank delivery-date criteria and includes additional party-group extraction procedures. The exact source columns and exclusions should be read from the VBA procedures before changing the business rules.

## 4. Transaction and account event handlers

LOCAL and NONLOCAL worksheet-change procedures respond to edited cells and perform party/miller lookups and related calculations. PARTY and MILLER event procedures support date-range or entity-based reporting. Results sheets contain selection-triggered report procedures.

## 5. Area-wise summaries

Individual worksheet-selection procedures are supplied for KOTHAPETA, PATAMATA, ELR, NZV, PNR, TNL, BHVM, MGR, BANDAR, GUNTUR, TIRUVURU, ELR ROAD, SINGHNAGAR, and SN PURAM. These are separate handlers in the source package, not a shared refactored implementation.

## Conceptual data flow

```text
PARTY / MILLER reference sheets
             |
             v
     LOCAL / NONLOCAL records
             |
      +------+-------+
      |              |
      v              v
Pending / regional   Party / miller
report extraction    account reports
      |              |
      +------+-------+
             v
     Output / results sheets
```

The diagram is conceptual. It does not represent a fully tested end-to-end process map.
