# VBA Project Structure - Superadmin Implementation

## Project Overview

This VBA project manages a validation workflow for pending changes between Admin and Superadmin files.

---

## File Structure

### Root Directory (Admin File: KindElternDaten_25_Admin.xlsm)

**Core Modules:**

- `DieseArbeitsmappe.cls` - Workbook events (Open, BeforeClose)
- `ExportSyncKartei.bas` - Main sync logic between Excel and Access
- `ExportUtilities.bas` - Dictionary operations, ID-based row updates, history tracking
- `ExportDataWork.bas` - DAO operations for tblKartei (read/write)
- `ExportProtection.bas` - Role-based permissions (Admin/Operator), past month validation
- `Notitzen.bas` - Mandatory comment mechanism for changes
- `ImportData.bas` - Import data from Access to Kartei sheet

**History & Risk Classification Modules:**

- `Export_HistoryBuilder.bas` - ✅ Extended history string builder with structured format
- `Export_HistoryConverter.bas` - ✅ Converts legacy history strings to new format (batch conversion for Kartei and Access tables)
- `Export_RiskClassification.bas` - ✅ Classifies changes as risky (Scenario A/B)
- `Export_OverlayPending.bas` - ✅ Overlays pre*/decl* records onto Kartei with color coding
- `Export_DeclinedTools.bas` - ✅ Tools for Admin to view/edit declined records (improved UX: positioned after Kartei, consistent headers/widths, last Decl_n comment only, AutoFilter, auto-cleanup)
- `Export_DeclinedHelpers.bas` - ✅ Helper functions for declined records management (comparison, copying, table creation)
- `Export_ManualImport.bas` - ✅ Manual database import (refresh Kartei from Access without reopening file)

**Utility Modules:**

- `AccessCreation.bas` - Database table creation
- `BaseBackupRestore.bas` - Backup/restore operations
- `DumpBase.bas` - Data export utilities
- `ExportBackupKartei.bas` - Kartei backup
- `ExportDeleteRecords.bas` - Record deletion
- `ExportIDAdd.bas` - ID management
- `ExportMessagePlanung.bas` - Scheduled messages
- `FormatCellsReal.bas` - Cell formatting
- `Kommentar.bas` - Comment extraction
- `SortName.bas` - Sorting utilities

---

### Superadmin Directory (Suprime File: KindElternDaten_25_Suprime.xlsm)

**Core Modules:**

- `valid_ImportPending.bas` - ✅ Load pending changes from pre_tblKartei into Kartei (with automatic monthly column formatting)
- `valid_ApproveFlow.bas` - ✅ Approve/Decline workflow and sync decisions (supports Mode A/B)
- `valid_ParseHistory.bas` - ✅ Tested history parser (from alt/Parceing.bas, DO NOT MODIFY REGEX)
- `valid_FormatMonths.bas` - ✅ Format monthly columns U-AF (21-32) to numeric with proper decimal separators
- `valid_TransposeForm.bas` - ✅ Transposed view for single record review (creates temp sheet geschichteForm)
- `Geschichte.bas` - ✅ Individual record history display (uses valid_ParseHistory)
- `grossGeschichte.bas` - ✅ Comprehensive history report generation (Mode A: all events, Mode B: last change per ID)

---

## Key Concepts

### 1. ID System (Column AV = 48)

- Every record has a unique numeric ID in column AV
- This ID matches exactly with `tblKartei.ID` in Access
- All dictionaries and operations use this ID as the primary key
- ID-based synchronization ensures correct record matching

### 2. Three-Table Architecture

```
tblKartei         → Main table (source of truth)
pre_tblKartei     → Pending risky changes (awaiting Superadmin approval)
decl_tblKartei    → Declined changes (rejected by Superadmin)
```

### 3. Risk Scenarios

**Scenario A (Non-strict):**

- AU ≠ "SEPA"
- Changes to past months (U-AF)
- No "NH" (case-sensitive) in J/O
- No "Nachhilfe", "Ind.", "VSpE" (case-insensitive) in J/O

**Scenario B (Strict):**

- AU = "SEPA"
- Any changes to months U-AF

### 4. Color Coding in Admin File

- **Light Blue (RGB 173, 216, 230)** → Pending records (from pre_tblKartei)
- **Red/Crimson (RGB 220, 20, 60)** → Declined records (from decl_tblKartei)
- Applied to column A if D ≠ "Zahlung"

### 5. History Format (Column AZ)

**New Structured Format (Export_HistoryBuilder):**

```
[RUCK:]<TAG>(<OLD>-><NEW>);<TAG>(<OLD>-><NEW>);.../@<COMMENT>@/<DATE>||
```

**Field Tags:**

| Tag          | Column       | Description                          |
| ------------ | ------------ | ------------------------------------ |
| `FID`        | A (1)        | Family ID                            |
| `PAR`        | B (2)        | Parent Name                          |
| `CHD`        | D (4)        | Child Name                           |
| `DOB`        | E (5)        | Date of Birth                        |
| `ADR`        | F (6)        | Address                              |
| `TEL`        | G (7)        | Phone                                |
| `MOB`        | H (8)        | Mobile                               |
| `EML`        | I (9)        | Email                                |
| `SB1`        | J (10)       | Subject (Months 1-6)                 |
| `PR1`        | M (13)       | Price (Months 1-6)                   |
| `SB2`        | O (15)       | Subject (Months 7-12)                |
| `PR2`        | R (18)       | Price (Months 7-12)                  |
| `M01`..`M12` | U-AF (21-32) | Monthly Charges                      |
| `EX1`        | AK (37)      | Extra Subject 1                      |
| `EX2`        | AL (38)      | Extra Subject 2                      |
| `EX3`        | AM (39)      | Extra Subject 3                      |
| `DCL`        | -            | Decline Entry (Superadmin rejection) |

**Delimiters:**

- `||` → Session separator (between change events)
- `;` → Field separator (within one session)
- `->` → Value separator (old value → new value)
- `/@` ... `@/` → Comment delimiters
- `RUCK:` → Prefix for retroactive (past month) changes

**Example:**

```
RUCK:M01(100.00->120.00);M02(80.00->90.00);ADR(Hauptstr. 1->Nebenstr. 2)/@Korrektur@/25.11.2025||FID(123->124);CHD(Hans->Franz)/@Fix@/20.11.2025||DCL(1->Record rejected [reason])||
```

**Backward Compatibility:**

Old format records (before migration) remain as-is. Parsers detect format by presence of `->`.

**Old Format (for reference):**

```
[Ruck: ]Mnt.N: War(X); Ist(Y). [FieldName: Was(X); Is(Y). ]/Comment/ dd.mm.yyyy ||
```

**History Conversion (Export_HistoryConverter):**

To convert existing legacy history strings to new format:

```vba
' Convert all histories in Kartei sheet
Call Export_HistoryConverter.ConvertAllHistoriesInKartei

' Convert histories in specific Access table
Call Export_HistoryConverter.ConvertHistoriesInAccessTable("tblKartei")

' Convert all (Kartei + all Access tables)
Call Export_HistoryConverter.ConvertAllHistoriesEverywhere

' Preview conversion without modifying data
Debug.Print Export_HistoryConverter.PreviewConversion("Mnt.8: War(12); Ist(24,6). /Comment/ 11.06.2025 ||")
```

Conversion mapping:

| Old Format                     | New Format                  |
| ------------------------------ | --------------------------- |
| `Mnt.8: War(12); Ist(24,6).`   | `M08(12->24,6)`             |
| `Address: Was(old); Is(new).`  | `ADR(old->new)`             |
| `Subject1: Was(old); Is(new).` | `SB1(old->new)`             |
| `Subject2: Was(old); Is(new).` | `SB2(old->new)`             |
| `Ruck: Mnt.9: War(X); Ist(Y).` | `RUCK:M09(X->Y)`            |
| `Decl_1: Was(); Is(comment).`  | `DCL(1->comment)`           |
| `/Comment/ 11.06.2025 \|\|`    | `/@Comment@/11.06.2025\|\|` |

---

## Workflow

### Admin Workflow

1. Opens file → Data loaded from tblKartei + overlays from pre*/decl*
2. Makes changes → Classified as risky or safe
3. **Safe changes** → Written directly to tblKartei
4. **Risky changes** → Sent to pre_tblKartei (pending approval)
5. **Declined records** → Can be edited and resubmitted (moved back to pre\_)

### Superadmin Workflow

1. Opens Suprime file
2. Runs `LoadPendingChanges` → Imports from pre_tblKartei, generates GrossGeschichte
3. **Chooses report mode:**
   - **Mode A:** Show all history events in date range
   - **Mode B:** Show only last change per ID in date range
4. Reviews changes in GrossGeschichte sheet
5. **Option A - Direct marking:** Marks each record in column AC (Decision):
   - `Approved` → Move to tblKartei
   - `Declined` → Move to decl_tblKartei (with comment in column AD)
6. **Option B - Transpose form:** Runs `ShowTransposeForm` for detailed single-record view:
   - Creates temporary sheet `geschichteForm` with vertical layout
   - Select Decision, optionally add Decline Comment
   - Click "Apply & Close" to save decision and return to GrossGeschichte
7. Runs `SyncDecisions` → Processes all decisions
   - In Mode A: If multiple blocks exist for same ID, only the decision from the last event is applied
   - In Mode B: Only one block per ID exists by design

### GrossGeschichte Column Structure (A-AE)

| Column | Field           | Format                    |
| ------ | --------------- | ------------------------- |
| A      | FamilyID        | Text                      |
| B      | Parent          | Text                      |
| C      | Child           | Text                      |
| D      | Birthdate       | Text (dd.mm.yyyy)         |
| E      | Address         | Text                      |
| F      | Phone           | Text                      |
| G      | Mobile          | Text                      |
| H      | Email           | Text                      |
| I      | Subject1        | Text                      |
| J      | Price1          | Text                      |
| K      | Subject2        | Text                      |
| L      | Price2          | Text                      |
| M-X    | Months 1-12     | Numeric (0.00)            |
| Y-AA   | Extra 1-3       | Text                      |
| AB     | Comments        | Text                      |
| AC     | Decision        | Approved/Declined         |
| AD     | Decline Comment | Text                      |
| AE     | RecordID        | Hidden (for internal use) |

---

## Technical Standards

### DAO Pattern

```vba
Dim engine As DAO.DBEngine
Set engine = New DAO.DBEngine
Dim wsDao As DAO.Workspace
Set wsDao = engine.Workspaces(0)
Dim db As DAO.Database
Set db = wsDao.OpenDatabase(dbPath)
```

### Performance

```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
' ... operations ...
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
```

### Error Handling

```vba
On Error GoTo ErrorHandler
' ... code ...
Exit Sub
ErrorHandler:
    ' Cleanup code
    MsgBox "Error: " & Err.Description
End Sub
```

### Naming Conventions

- **Admin modules:** `Export_*` prefix
- **Superadmin modules:** `valid_*` prefix
- **English:** All comments, MsgBox texts, UI elements

---

### Database Tables

### tblKartei Structure

```
ID               → AutoNumber (primary key)
Value1..Value51  → Text(255) [data columns A-AY], AllowZeroLength = True
Value52          → Memo [history column AZ]
InteriorColor1..InteriorColor51 → Long [cell background colors]
FontColor3       → Long [font color for column C]
FontColor18      → Long [font color for column R]
```

### pre_tblKartei Structure

Same as tblKartei, but ID is regular Long (not AutoNumber) - preserves original ID from tblKartei.

### decl_tblKartei Structure

Same as tblKartei, but ID is regular Long (not AutoNumber) - preserves original ID from tblKartei.
**Important:** All text fields must have `AllowZeroLength = True` to accept empty strings.

---

## Next Steps

### For Admin File:

- ✅ Implement `Export_RiskClassification.bas` (classify changes)
- ✅ Implement `Export_OverlayPending.bas` (overlay pre*/decl* on Kartei)
- ✅ Implement `Export_DeclinedTools.bas` (handle declined records)
- [ ] Integrate risk classification into `CompareAndSyncKartei`
- [ ] Block SEPA changes for Operator role
- [ ] Integrate overlay logic into `Workbook_Open`

### For Superadmin File:

- ✅ `valid_ImportPending.bas` - Load pending changes
- ✅ `valid_ApproveFlow.bas` - Approve/decline workflow
- ✅ `valid_ParseHistory.bas` - History parser
- ✅ `Geschichte.bas` - Individual record history
- ✅ `grossGeschichte.bas` - Comprehensive history report

---

## Testing Checklist

### Admin Side:

- [ ] Open file → pre*/decl* records overlay correctly
- [ ] Safe changes → Write directly to tblKartei
- [ ] Risky changes → Write to pre_tblKartei
- [ ] Declined record edit → Moves back to pre_tblKartei
- [ ] Pending records → Cannot be written to tblKartei directly
- [ ] SEPA records → Operator cannot modify

### Superadmin Side:

- [x] Load pending → Kartei populated from pre_tblKartei
- [x] GrossGeschichte Mode A → Shows all events per ID in date range
- [x] GrossGeschichte Mode B → Shows only last change per ID in date range
- [x] Mode B filtering → Correctly identifies latest date and last event in collection
- [x] Approve decision → Moves to tblKartei, deletes from pre\_
- [x] Decline decision → Moves to decl*tblKartei with comment, deletes from pre*
- [x] Conflicting decisions (Mode A) → Only decision from last event is applied per ID
- [x] History parsing → Supports months, Address, Subject1/2, Decl_N
- [x] decl_tblKartei creation → AllowZeroLength = True for text fields

---

## Contact & Maintenance

- All code is modular and documented in English
- Each function/procedure has a clear, single responsibility
- No monolithic procedures
- Use existing DAO patterns consistently
