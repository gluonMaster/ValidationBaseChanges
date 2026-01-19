# Superadmin Multi-Year User Guide

Central administration workbook for years 2024, 2025, and 2026.

## Overview

The Superadmin workbook (`KindElternDaten_Suprime.xlsm`) is a centralized tool for reviewing and approving pending changes across three year databases:

- `KindElternDaten_24_front.accdb` (Year 2024)
- `KindElternDaten_25_front.accdb` (Year 2025)
- `KindElternDaten_26_front.accdb` (Year 2026)

Instead of opening separate yearly workbooks, all operations are performed in one workbook with year-specific sheets.

## Sheet Names

### Per-Year Sheets

| Purpose                                    | Year 2024           | Year 2025           | Year 2026           |
| ------------------------------------------ | ------------------- | ------------------- | ------------------- |
| Pending data (from `pre_tblKartei`)        | `Kartei24`          | `Kartei25`          | `Kartei26`          |
| Decision sheet (War/Ist + Approve/Decline) | `grossGeschichte24` | `grossGeschichte25` | `grossGeschichte26` |
| Single-record history                      | `Geschichte24`      | `Geschichte25`      | `Geschichte26`      |
| Date-range history (all records)           | `Geschichte24_Alle` | `Geschichte25_Alle` | `Geschichte26_Alle` |

### Shared Sheets

| Sheet       | Purpose                                     |
| ----------- | ------------------------------------------- |
| `Dashboard` | Year selection, quick actions, instructions |
| `Neu`       | New records list (only for 2026)            |
| `DBConfig`  | Database path configuration (hidden)        |
| `NeuConfig` | LastSeenID storage (hidden)                 |

## Database Configuration

### Initial Setup

1. Open the workbook for the first time
2. Run `Dashboard_ConfigureDatabase` (or `ConfigureAllDatabases`)
3. Select the cloud root folder that contains `2024`, `2025`, and `2026` subfolders
4. The system expects this structure:

```
<CloudRoot>/
  2024/
    Alarm/
      KindElternDaten_24_front.accdb
  2025/
    Alarm/
      KindElternDaten_25_front.accdb
  2026/
    Alarm/
      KindElternDaten_26_front.accdb
```

5. If a year is in a different location, you can set per-year overrides via:
   - `SelectDbFolder24`
   - `SelectDbFolder25`
   - `SelectDbFolder26`

### Viewing Current Configuration

Run `Dashboard_ShowConfig` to see:

- Cloud root path (if set)
- Year overrides (if set)
- Expected DB path per year + status (`ok`, `missing-file`, `not-configured`)

Note: `Dashboard_ShowConfig` is non-interactive (it should not open folder pickers).

## Workflow: Approve/Decline Pending Changes

### Step 1: Select Year

Use Dashboard buttons or macros:

```vba
SelectYear24   ' or SelectYear25, SelectYear26
```

Or set the year in `Dashboard!B2` manually (24, 25, or 26).

### Step 2: Load Pending Changes (+ Decision Sheet)

```vba
Dashboard_LoadPending
```

This will:

1. Load pending records from `pre_tblKartei` into `KarteiYY`
2. Build the decision sheet `grossGeschichteYY` with War/Ist comparisons
3. Add dropdowns for Approve/Decline in decision columns

If you only need to rebuild the decision sheet (without re-importing pending), use:

```vba
Dashboard_BuildDecisionSheet
```

### Step 3: Review Changes

1. Go to `grossGeschichteYY` sheet
2. Each pending record has two rows:
   - War row: Original values from `tblKartei` (or empty for new records)
   - Ist row: Pending values from `pre_tblKartei`
3. Differences are highlighted
4. Set decision in the decision column:
   - `Approved` -> move to `tblKartei`
   - `Declined` -> move to `decl_tblKartei`
5. For declined entries, add an optional comment in the comment column

### Step 4: Sync Decisions

```vba
Dashboard_SyncDecisions
```

This writes decisions back to the Access database for the selected year.

## Workflow: History Reports

### Single-Record History

1. Go to `KarteiYY` sheet
2. Select any cell in the row you want to analyze
3. Run:

```vba
Dashboard_SingleRecordHistory
```

Output appears in `GeschichteYY` sheet.

### Date-Range History (All Records)

1. Set date range in `GeschichteYY_Alle!B1` (start) and `C1` (end)
2. Run:

```vba
Dashboard_DateRangeHistory
```

Depending on implementation settings, the macro may offer a choice between:
- building history from `KarteiYY` (pending subset), or
- building history directly from the year database (`tblKartei`) (full DB mode)

Output appears in `GeschichteYY_Alle`.

## Workflow: Payment Export (FamZahlungen)

```vba
Dashboard_FamZahlungen
```

Exports files to `C:\FamZahlung` with year-suffixed filenames (e.g. `EltKosten_25.xlsx`).

Note: payment exports work on `KarteiYY` sheet data. For complete reports, ensure the sheet contains full `tblKartei` data (not only pending from `pre_tblKartei`).

## Workflow: New Records List (2026 Only)

```vba
Dashboard_RefreshNeu   ' or RefreshNeuList
```

This refreshes the `Neu` sheet with records where `ID > LastSeenID`.
Only relevant for year 2026 (no new records expected for past years).

## Macro Reference

### Dashboard Macros (Use Selected Year)

| Macro                            | Description                                 |
| -------------------------------- | ------------------------------------------- |
| `SelectYear24/25/26`             | Set the active year for Dashboard actions   |
| `Dashboard_LoadPending`          | Load pending + build decision sheet         |
| `Dashboard_BuildDecisionSheet`   | Rebuild decision sheet (no pending import)  |
| `Dashboard_SyncDecisions`        | Sync approved/declined decisions to DB      |
| `Dashboard_SingleRecordHistory`  | Build history for selected row              |
| `Dashboard_DateRangeHistory`     | Build history for date range                |
| `Dashboard_FamZahlungen`         | Export payment files                        |
| `Dashboard_RefreshNeu`           | Refresh new records list (2026)             |
| `Dashboard_ConfigureDatabase`    | Configure database paths                    |
| `Dashboard_ShowConfig`           | Show current configuration (no prompting)   |
| `ShowDashboard`                  | Activate/create the Dashboard worksheet     |
| `QuickAction`                    | InputBox: select year + action              |

### Direct Year Macros

| Action                  | Year 24                         | Year 25                         | Year 26                         |
| ----------------------- | ------------------------------- | ------------------------------- | ------------------------------- |
| Load pending + decision | `LoadPendingAndBuildDecision24` | `LoadPendingAndBuildDecision25` | `LoadPendingAndBuildDecision26` |
| Build decision only     | `BuildPendingDecisionSheet24`   | `BuildPendingDecisionSheet25`   | `BuildPendingDecisionSheet26`   |
| Sync decisions          | `SyncDecisions24`               | `SyncDecisions25`               | `SyncDecisions26`               |
| Single-record history   | `GeschichteMachen24`            | `GeschichteMachen25`            | `GeschichteMachen26`            |
| Date-range history      | `GrossGeschichteMachen24`       | `GrossGeschichteMachen25`       | `GrossGeschichteMachen26`       |
| Payment export          | `FamZahlungen24`                | `FamZahlungen25`                | `FamZahlungen26`                |
| Configure DB path       | `SelectDbFolder24`              | `SelectDbFolder25`              | `SelectDbFolder26`              |

## Workbook Behavior

### On Open

- Shows/creates the `Dashboard` sheet
- Optionally auto-refreshes `Neu` (2026), but avoids prompting if the DB path is not yet configured

### On Close

- Clears transient data from `KarteiYY` and `grossGeschichteYY` sheets (rows 3+)
- Preserves:
  - Dashboard configuration
  - `DBConfig` (database paths)
  - `NeuConfig` (LastSeenID)
  - History sheets (`GeschichteYY`, `GeschichteYY_Alle`)
  - `Neu` list
- Auto-saves without prompting

### Disabling Auto-Clear

If you need to preserve Kartei/grossGeschichte data across sessions (e.g., for debugging), edit `Superadmin/DieseArbeitsmappe.cls`:

```vba
Private Const CLEAR_ON_CLOSE As Boolean = False  ' Change from True to False
```

## Module Reference

| Module                              | Purpose                                 |
| ----------------------------------- | --------------------------------------- |
| `Superadmin/valid_Dashboard.bas`    | Dashboard UI helpers and year selection |
| `Superadmin/valid_YearConfig.bas`   | Multi-year DB configuration             |
| `Superadmin/valid_ApproveFlow.bas`  | Load pending + sync decisions           |
| `Superadmin/valid_ImportPending.bas`| Import from pre_tblKartei               |
| `Superadmin/valid_GrossGeschichteDecision.bas` | Build War/Ist decision sheet    |
| `Superadmin/valid_HistoryPerYear.bas` | History tools per year                |
| `Superadmin/pay_Main.bas`           | Payment export entry points             |
| `Superadmin/valid_NeuList.bas`      | New records list (2026)                 |
| `Superadmin/DieseArbeitsmappe.cls`  | Workbook events (open/close)            |

_Last updated: January 2026_

