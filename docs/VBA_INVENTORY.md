# VBA Source Inventory

## Standard modules

| File | Procedures / purpose |
|---|---|
| `modCommissionValidation.bas` | `CheckAllCommissionRates`, `ValidateSheetRates`; checks LOCAL and NONLOCAL commission rates against the PARTY master list. |
| `modPrakasamReports.bas` | `IsPrakasamParty`, `ClearPrakasamTables`, `Extract_Prakasham_Pending_Delivery`, `Extract_Prakasham_Pending_Commission`; helper routines and Prakasam-specific reports from NONLOCAL data. |
| `modPendingDelivery.bas` | `UpdatePendingLocal`, `UpdatePendingNonLocal`, `ClearAutomationOutput`, `RefreshPendingDelivery`, `ExtractByPartyPrefixes`, `Extract_KALI`, `Extract_PRASAD`, `Extract_MADHU`, `Extract_JEMILI`, `Extract_MK`, `Extract_ONG_VRR`; pending-delivery refresh and party-group extraction procedures. |

## Worksheet event code

| Source file | Workbook worksheet |
|---|---|
| `Sheet_LOCAL.cls` | `LOCAL` |
| `Sheet_NONLOCAL.cls` | `NONLOCAL` |
| `Sheet_PARTY.cls` | `PARTY` |
| `Sheet_MILLER.cls` | `MILLER` |
| `Sheet_Results_LOCAL.cls` | `Results_LOCAL` |
| `Sheet_Results_NONLOCAL.cls` | `Results_NONLOCAL` |
| `Sheet_MILLER_LOCAL.cls` | `MILLER_LOCAL` |
| `Sheet_Results_MILLER_LOCAL.cls` | `Results_MILLER_LOCAL` |
| `Sheet_PENDING_DELIVERY.cls` | `PENDING_DELIVERY` |
| `Sheet_PRAKASAM.cls` | `PRAKASAM` |
| `Sheet_KOTHAPETA.cls` | `KOTHAPETA` |
| `Sheet_PATAMATA.cls` | `PATAMATA` |
| `Sheet_ELR.cls` | `ELR` |
| `Sheet_NZV.cls` | `NZV` |
| `Sheet_PNR.cls` | `PNR` |
| `Sheet_TNL.cls` | `TNL` |
| `Sheet_BHVM.cls` | `BHVM` |
| `Sheet_MGR.cls` | `MGR` |
| `Sheet_BANDAR.cls` | `BANDAR` |
| `Sheet_GUNTUR.cls` | `GUNTUR` |
| `Sheet_TIRUVURU.cls` | `TIRUVURU` |
| `Sheet_ELR_ROAD.cls` | `ELR ROAD` |
| `Sheet_SINGHNAGAR.cls` | `SINGHNAGAR` |
| `Sheet_SN_PURAM.cls` | `SN PURAM` |

## Important implementation note

The worksheet `.cls` files are provided as text source snapshots for the matching worksheet code windows. They are not intended to be imported as standalone class modules. See the import/paste instructions in the root README.

The current package deliberately does not claim that the code was refactored, that a specific number of lines was removed, or that application-state restoration was standardized. No such change has been made as part of this package rebuild.
