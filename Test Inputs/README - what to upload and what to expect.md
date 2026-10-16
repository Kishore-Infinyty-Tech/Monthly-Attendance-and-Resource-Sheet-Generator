# Test input files (same format as the client's files)

Target month is **July 2026** (31 days). Resource sheets are for **June 2026** (the previous month).

## Main set - upload these together
| Upload as | File |
|---|---|
| Master | `01_Master_Sangraha_July-2026.xlsx` (19 employees, tabs `Monthly_July_Infinity` + `Daily Attendance_Infinity`) |
| Resource sheet | `02_Resource_Sheet_ALL_Employees_June.xlsx` (18 employees in one sheet) |

Expected: no errors, 5 warnings, 18 employees, 37 files in the ZIP. Zoya Khan is in the master only and is ignored.
The combined sheet has 18 employee rows, so its Total row moves down from row 19 to row 23.

| Employee | Case covered | Expected |
|---|---|---|
| Rahul Desai | Fully present, no leave | 31 of 31 days, 100,000, nothing red except PO |
| Richa B | Leave inside balance, different rate (85,000) | Leave till date 2 + 5 = 7 |
| Aakash Singh | Rate 72,000 | Amount 72,000 |
| Rahul Singh | Holiday day | Holiday counted as billed |
| Raman Rao | Many absent days | 8 of 31 days, amount **red** |
| Brijesh Yadav | Absent all month | 0 days, amount 0, **red** |
| Maitri Kaur | Leave + absent, no week off | 14 of 31 days, amount **red** |
| Mahesh Patil | Almost always absent | 9 of 31 days, amount **red** |
| Meena Jain | Leave over the allowed total | 23 + 2 = 25 > 24, **whole row red** |
| Sneha Rao | Leave exactly equal to total | 20 + 4 = 24, **not red** |
| Vikram Shah | CompOff days, Total Leave 18, rate 110,000 | Warning about CompOff, 29 billing days, amount red |
| Priya Nair | Half-day leave (0.5), rate 91,000 | Leave till date 2.5 |
| Amit Kumar (x2) | Same name, different IDs | Both matched correctly (by ID tie-break) |
| Karan Mehta | Name typo "Karan Mehra" in resource sheet | Warning, matched by Vendor Resource ID |
| Divya Pillai | Name "Divya P." and VMT id in Vendor ID column | Warning, matched by Vendor ID |
| Sanjay Gupta | Day counts add up to 28, master total says 31 | Warning, 28 billing days |
| Anita Desai | "  anita   DESAI " (spaces/capitals) | Matches by name, no warning |

## Individual resource sheets - `individual/`
One file per employee (Rahul, Richa, Maitri, Meena, Sneha, Karan, Priya, Vikram).
Upload the master + all of these together (multi-select). Expected: no errors, 3 warnings (Karan, Meena, Vikram), 8 employees.

## Edge cases - `edge_cases/` (use with the main master unless stated)
| File | Expected |
|---|---|
| `E1_master_renamed_headers_and_tabs.xlsx` (as master, with an individual resource sheet) | Works: headers renamed, tab `monthly - July`, extra `Weekly` tab, 2 title rows above the header |
| `E2_resource_renamed_headers_and_tab.xlsx` | Works: headers like "Employee Name", "Rate / Month", "PO No", tab renamed |
| `E3_resource_employee_not_in_master_ERROR.xlsx` | **Error**, no download |
| `E4a_...duplicate_employee_A.xlsx` + `E4b_...duplicate_employee_B.xlsx` (upload both) | Warning: Rahul Desai duplicate, first one used |
| `E5_resource_sheet_already_July_WARNING.xlsx` | Warning: sheet is already July |
| `E6_resource_missing_rate_column_ERROR.xlsx` | **Error**: Rate/Month column not found |
| `E7_master_tab_not_named_monthly.xlsx` (as master, with an individual resource sheet) | Warning: tab not named Monthly, used by columns |
| `E8_master_month_cell_blank.xlsx` (as master, with an individual resource sheet) | **Error** until you pick July 2026 in the Month box |
| `E9_resource_sheet_older_month_WARNING.xlsx` | Warning: sheet is December 2025, not June 2026 |
