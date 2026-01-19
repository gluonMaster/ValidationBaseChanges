# Merge Migration Map: Superadmin Multi-Year + Data-file Port

> Generated from discovery of `Superadmin/` and `Data-file/` modules.  
> See `PROMPTS/00_AGENT_CONTEXT.md` for full requirements.

---

## 1. Current Superadmin Implementation

### 1.1 Main Macros (Entry Points)

| Macro                       | Module                            | Purpose                                                                                                     |
| --------------------------- | --------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `LoadPendingChanges`        | valid_ApproveFlow.bas             | Main workflow: loads pre_tblKartei → Kartei, builds GrossGeschichte decision sheet, adds decision dropdowns |
| `SyncDecisions`             | valid_ApproveFlow.bas             | Processes Approved/Declined decisions from GrossGeschichte sheet                                            |
| `LoadPendingChangesFromPre` | valid_ImportPending.bas           | Raw import from `pre_tblKartei` into `Kartei` sheet                                                         |
| `BuildPendingDecisionSheet` | valid_GrossGeschichteDecision.bas | Creates War/Ist comparison for pending records on `grossGeschichte`                                         |
| `RefreshNeuList`            | valid_NeuList.bas                 | Refreshes `Neu` sheet with new records (ID > LastSeenID)                                                    |
| `GeshichteMachen`           | Geschichte.bas                    | Single-record history report → `Geschichte` sheet                                                           |
| `GrossGeshichteMachen`      | grossGeschichte.bas               | Date-range history (view-only) → `Geschichte` sheet                                                         |
| `SelectDatabaseFolder`      | valid_DatabasePath.bas            | Manual DB folder selector                                                                                   |
| `ShowTransposeForm`         | valid_TransposeForm.bas           | Transpose view of GrossGeschichte entry                                                                     |
| `FormatMonthlyColumns`      | valid_FormatMonths.bas            | Formats monthly columns U-AF                                                                                |

### 1.2 Hard-Coded Sheet Names

| Module                            | Sheet Reference   | Purpose                                |
| --------------------------------- | ----------------- | -------------------------------------- |
| valid_DatabasePath.bas            | `Kartei`          | Read/write DB base path in cell **X1** |
| valid_ImportPending.bas           | `Kartei`          | Write pending data from pre_tblKartei  |
| valid_ImportPending.bas           | `GrossGeschichte` | Set date range in B1/C1                |
| valid_ApproveFlow.bas             | `grossGeschichte` | Decision column dropdowns (AC/AD)      |
| valid_GrossGeschichteDecision.bas | `grossGeschichte` | War/Ist decision sheet output          |
| valid_GrossGeschichteDecision.bas | `Kartei`          | Source of pending records              |
| valid_GrossGeschichteData.bas     | `Kartei`          | Read pending "Ist" values              |
| valid_NeuList.bas                 | `Neu`             | Output sheet for new records list      |
| valid_NeuList.bas                 | `NeuConfig`       | Hidden sheet storing LastSeenID        |
| valid_NeuList.bas                 | `Kartei`          | Copy header row to Neu                 |
| Geschichte.bas                    | `Kartei`          | Source for single-record history       |
| Geschichte.bas                    | `Geschichte`      | Output sheet                           |
| grossGeschichte.bas               | `Kartei`          | Source data                            |
| grossGeschichte.bas               | `Geschichte`      | Output sheet (view-only)               |
| valid_TransposeForm.bas           | `GrossGeschichte` | Source for transpose                   |
| valid_TransposeForm.bas           | `geschichteForm`  | Temp output sheet                      |
| valid_FormatMonths.bas            | `Kartei`          | Format monthly columns                 |

### 1.3 Database Path Logic

| Module                        | Storage Location                                      | DB File Constant                   |
| ----------------------------- | ----------------------------------------------------- | ---------------------------------- |
| valid_DatabasePath.bas        | Hidden `DBConfig` sheet (A1 root, B1/C1/D1 overrides) | `KindElternDaten_<YY>_front.accdb` |
| valid_DatabasePath.bas        | Subfolder constant: `Alarm`                           | —                                  |
| valid_GrossGeschichteData.bas | Delegates to `valid_DatabasePath`                     | —                                  |
| valid_ImportPending.bas       | Uses `GetDatabasePath()` → `valid_DatabasePath`       | —                                  |
| valid_NeuList.bas             | Uses `GetDatabasePath()` → `valid_DatabasePath`       | —                                  |

**Key points:**

- DB filename is year-dynamic: `KindElternDaten_<YY>_front.accdb`
- Cloud root stored in `DBConfig!A1` (user selects folder containing `2024/2025/2026`)
- Full path (default) = `<Root>\<YYYY>\Alarm\KindElternDaten_<YY>_front.accdb` (per-year override in `DBConfig!B1/C1/D1` if needed)
- User is prompted via folder/file dialog if a DB is missing

---

## 2. Data-file Reference Modules (to Port)

### 2.1 Single-Record History

| Module         | Entry Point        | Output Sheet        | Dependencies                     |
| -------------- | ------------------ | ------------------- | -------------------------------- |
| Geschichte.bas | `GeschichteMachen` | `Geschichte_Einzel` | `History_ParseData.ParseHistory` |

**Notes:**

- Reads history from `Kartei!AZ` (column 52)
- Outputs to `Geschichte_Einzel` sheet (different name than Superadmin's `Geschichte`)
- Uses `History_ParseData.ParseHistory` (same interface as Superadmin `valid_ParseHistory.ParseHistory`; pick one to avoid duplication)

### 2.2 Date-Range History (All Records)

| Module              | Entry Point            | Output Sheet      | Dependencies                     |
| ------------------- | ---------------------- | ----------------- | -------------------------------- |
| grossGeschichte.bas | `GrossGeshichteMachen` | `Geschichte_Alle` | `History_ParseData.ParseHistory` |

**Notes:**

- Two modes: A = all events, B = last change per ID
- **View-only** (no decision columns AC/AD/AE)
- Date range from `B1:C1`
- Source: `Kartei` sheet

### 2.3 Import from Access

| Module         | Entry Point      | Source Table | Output Sheet |
| -------------- | ---------------- | ------------ | ------------ |
| ImportData.bas | `ImportFromBase` | `tblKartei`  | `Kartei`     |

**Notes:**

- Loads **full** `tblKartei` (not pre_tblKartei)
- DB path read from `Kartei!I1` (different cell than Superadmin's X1)
- Applies formatting (interior colors, font colors)
- Currently hard-coded for year 2025: `KindElternDaten_25_front.accdb`

### 2.4 Payment Exports

| Module                 | Entry Point    | Output Files                                                            | Source Sheet |
| ---------------------- | -------------- | ----------------------------------------------------------------------- | ------------ |
| zahl_Main.bas          | `FamZahlungen` | `C:\FamZahlung\EltKosten.xlsx`, `EltKostenNH.xlsx`, `EltKostenInd.xlsx` | `Kartei`     |
| zahl_DataProcessor.bas | (internal)     | —                                                                       | `Kartei`     |
| zahl_FileGenerator.bas | (internal)     | —                                                                       | —            |
| zahl_Utils.bas         | (internal)     | —                                                                       | —            |

**Notes:**

- Aggregates payments by family ID across all children
- Three output files: Regular, Nachhilfe (tutoring), Individual
- Output folder: `C:\FamZahlung` (hard-coded)
- Output filenames not year-suffixed (will overwrite if run multiple times)
- In Superadmin, this functionality is ported as `Superadmin/pay_*.bas` with year-suffixed output filenames.

---

## 3. Migration Map Table

| Feature                                | Current Module(s)                   | Source Sheet(s)           | Target Per-Year Sheet(s)                | DB Table(s)                              | Notes / Refactor Points                                                   |
| -------------------------------------- | ----------------------------------- | ------------------------- | --------------------------------------- | ---------------------------------------- | ------------------------------------------------------------------------- |
| **Import pending**                     | `valid_ImportPending.bas`           | Kartei                    | Kartei24, Kartei25, Kartei26            | pre_tblKartei                            | Add year param; update sheet refs                                         |
| **Decision sheet build**               | `valid_GrossGeschichteDecision.bas` | Kartei, grossGeschichte   | grossGeschichte24/25/26                 | tblKartei, pre_tblKartei                 | Add year param; sheet name suffix                                         |
| **Approve/Decline sync**               | `valid_ApproveFlow.bas`             | grossGeschichte, Kartei   | grossGeschichte24/25/26, Kartei24/25/26 | pre_tblKartei, tblKartei, decl_tblKartei | Year param for all sync ops                                               |
| **Neu list**                           | `valid_NeuList.bas`                 | Neu, Kartei               | **Neu (2026 only)**                     | tblKartei                                | Only needed for 2026; keep sheet name `Neu`, but force DB=2026            |
| **DB path config**                     | `valid_DatabasePath.bas`            | Kartei!X1 (legacy)        | **DBConfig sheet (hidden)**             | —                                        | **DONE** - uses valid_YearConfig.bas with cloud root + per-year overrides |
| **Single-record history**              | `valid_HistoryPerYear.bas`          | KarteiYY                  | Geschichte24/25/26                      | —                                        | **DONE** - GeschichteMachenForYear(year2) + wrapper macros                |
| **Date-range history**                 | `valid_HistoryPerYear.bas`          | tblKartei (DB) / KarteiYY | Geschichte24_Alle / 25_Alle / 26_Alle   | —                                        | **DONE** - GrossGeschichteMachenForYear(year2) supports DB or KarteiYY    |
| **PORT: Single-record history (Data)** | `valid_HistoryPerYear.bas`          | KarteiYY                  | GeschichteYY                            | —                                        | **DONE** - merged into valid_HistoryPerYear                               |
| **PORT: Date-range history (Data)**    | `valid_HistoryPerYear.bas`          | KarteiYY                  | GeschichteYY_Alle                       | —                                        | **DONE** - merged into valid_HistoryPerYear                               |
| **PORT: Import full Kartei**           | `Data-file/ImportData.bas`          | Kartei                    | Kartei24/25/26                          | tblKartei                                | If needed for FamZahlungen on full data                                   |
| **PORT: Payment exports**              | `pay_Main.bas`, `pay_*.bas`         | Kartei24/25/26            | Kartei24/25/26                          | —                                        | **DONE** - FamZahlungenForYear(year2) + wrappers; year-suffixed output    |
| **History parser**                     | `valid_ParseHistory.bas`            | —                         | —                                       | —                                        | **SAFE** – no sheet refs; shared by all                                   |
| **Format months**                      | `valid_FormatMonths.bas`            | Kartei                    | Kartei24/25/26                          | —                                        | Add sheet param                                                           |
| **Transpose form**                     | `valid_TransposeForm.bas`           | GrossGeschichte           | geschichteForm                          | —                                        | Update to detect year from source sheet                                   |
| **Data access layer**                  | `valid_GrossGeschichteData.bas`     | —                         | —                                       | tblKartei                                | Add year param for DB selection                                           |

---

## 4. Modules Safe to Leave Untouched

| Module                            | Reason                                                                                    |
| --------------------------------- | ----------------------------------------------------------------------------------------- |
| `valid_ParseHistory.bas`          | Self-contained parser; no sheet or DB dependencies                                        |
| `Data-file/History_ParseData.bas` | Duplicate of valid_ParseHistory; may not need to port                                     |
| `Data-file/Parceing.bas`          | Older parser; superseded by History_ParseData                                             |
| `Data-file/SortName.bas`          | Utility; not needed for the merge (hard-codes `Kartei`)                                   |
| `Data-file/FormatCellsReal.bas`   | Utility; works on `ActiveSheet` (activate target sheet or refactor to accept `Worksheet`) |

---

## 5. High-Level Refactor Strategy

### 5.1 Database Path Management

1. Extend `valid_DatabasePath.bas` to manage **three** DB paths (2024/2025/2026)
2. Use dedicated hidden config sheet: `DBConfig`:
   - `A1`: Cloud root folder (contains `2024/2025/2026`)
   - `B1/C1/D1`: per-year folder overrides (optional)
3. Add `GetValidatedDatabasePathForYear(year As Integer)` function
4. Update `DB_FILENAME` constant to be year-dynamic: `KindElternDaten_<YY>_front.accdb`
5. Cloud setup (user-confirmed): prefer selecting ONE folder that contains `2024/2025/2026`, then derive
   `<Root>\\2024\\Alarm\\KindElternDaten_24_front.accdb`, etc. Fall back to per-year overrides only if needed.

### 5.2 Sheet Naming Convention

- Pending sheets: `Kartei24`, `Kartei25`, `Kartei26`
- Decision sheets: `grossGeschichte24`, `grossGeschichte25`, `grossGeschichte26`
- History single: `Geschichte24`, `Geschichte25`, `Geschichte26`
- History all: `Geschichte24_Alle`, `Geschichte25_Alle`, `Geschichte26_Alle`
- Neu list: `Neu` (only year 2026)

### 5.3 Parameterized Helpers Pattern

```vba
' Example: Parameterized import
Public Sub LoadPendingChangesForYear(year As Integer)
    Dim sheetName As String
    sheetName = "Kartei" & Right(CStr(year), 2)
    ' ... rest of logic using sheetName
End Sub

' Thin wrappers for buttons/menus
Public Sub LoadPending24(): LoadPendingChangesForYear 2024: End Sub
Public Sub LoadPending25(): LoadPendingChangesForYear 2025: End Sub
Public Sub LoadPending26(): LoadPendingChangesForYear 2026: End Sub
```

### 5.4 Payment Export Porting

1. [done] Copy `zahl_*.bas` modules to `Superadmin/`
2. [done] Rename to `pay_*.bas` to avoid confusion
3. [done] Add year parameter to `FamZahlungen`:
   - Accept sheet name param (e.g. `Kartei25`) or `year2`
   - Suffix output files: `EltKosten_25.xlsx`
4. [done] Create thin wrappers: `FamZahlungen24`, `FamZahlungen25`, `FamZahlungen26`

### 5.5 History Tools Porting

1. Superadmin already has `Geschichte.bas` and `grossGeschichte.bas`
2. Merge Data-file versions if they offer features Superadmin lacks:
   - Data-file outputs to `Geschichte_Einzel` / `Geschichte_Alle`
   - Superadmin outputs to `Geschichte`
3. Add year parameter and update output sheet names

---

## 6. Open Questions (Before Implementation)

1. **Backwards compatibility:** Should we keep non-suffixed sheet names (`Kartei`, `grossGeschichte`) as aliases for 2025, or fully migrate to per-year names?

2. **Data source for FamZahlungen:** Current Data-file loads from `tblKartei` (full/approved data). Should Superadmin's `KarteiYY` sheets contain:
   - Pending-only (from `pre_tblKartei`) – current behavior
   - Full data (from `tblKartei`) – needed for FamZahlungen
   - Or add separate "full import" button per year?

3. **Payment output paths:** Should `C:\FamZahlung\EltKosten.xlsx` become:
   - `C:\FamZahlung\EltKosten_24.xlsx` (year suffix)
   - `C:\FamZahlung\2024\EltKosten.xlsx` (year subfolder)
   - User-configurable?

4. **NeuConfig storage:** Neu is only needed for 2026, so `NeuConfig` can stay single-year (2026).

---

## 7. File Reference Quick Lookup

### Superadmin/ (Target)

```
DieseArbeitsmappe.cls         – Workbook events (multi-year safe: clears KarteiYY/grossGeschichteYY on close)
Geschichte.bas                – Single-record history (auto-detects year, routes to valid_HistoryPerYear)
grossGeschichte.bas           – Date-range history (auto-detects year, routes to valid_HistoryPerYear)
valid_ApproveFlow.bas         – Main workflow + decision sync
valid_Dashboard.bas           – [NEW] Dashboard UI helpers (year selection, action dispatchers, quick access)
valid_DatabasePath.bas        – DB path management (delegates to valid_YearConfig)
valid_FormatMonths.bas        – Format monthly columns
valid_GrossGeschichteData.bas – Data access layer
valid_GrossGeschichteDecision.bas – War/Ist decision sheet builder
valid_HistoryPerYear.bas      – [NEW] Year-aware history: GeschichteMachenForYear, GrossGeschichteMachenForYear
valid_ImportPending.bas       – Import from pre_tblKartei
valid_NeuList.bas             – New records list
valid_ParseHistory.bas        – History string parser (SAFE)
valid_TransposeForm.bas       – Transpose view helper
valid_YearConfig.bas          – [NEW] Multi-year DB config (cloud root, per-year overrides, sheet name helpers)
pay_Main.bas                  – [NEW] Payment export main: FamZahlungenForYear, FamZahlungen24/25/26
pay_DataProcessor.bas         – [NEW] Payment data processing (sheet-agnostic)
pay_FileGenerator.bas         – [NEW] Payment file generation (year-suffixed filenames)
pay_Utils.bas                 – [NEW] Payment utilities (sheet-agnostic)
```

### Data-file/ (Source for Porting)

```
Geschichte.bas                – [PORTED] single-record history → valid_HistoryPerYear.bas
grossGeschichte.bas           – [PORTED] date-range history → valid_HistoryPerYear.bas
ImportData.bas                – Port: full import from tblKartei
History_ParseData.bas         – Parser (duplicate of valid_ParseHistory)
zahl_Main.bas                 – [PORTED] payment export main → pay_Main.bas
zahl_DataProcessor.bas        – [PORTED] payment processing → pay_DataProcessor.bas
zahl_FileGenerator.bas        – [PORTED] payment file generation → pay_FileGenerator.bas
zahl_Utils.bas                – [PORTED] payment utilities → pay_Utils.bas
FormatCellsReal.bas           – Utility (number/date normalization; currently in repo root/Data-file)
SortName.bas                  – Utility
Parceing.bas                  – Legacy parser (skip)
```

---

_End of Migration Map_
