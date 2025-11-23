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

**Risk Classification Modules:**

- `Export_RiskClassification.bas` - ✅ Classifies changes as risky (Scenario A/B)
- `Export_OverlayPending.bas` - ✅ Overlays pre*/decl* records onto Kartei with color coding
- `Export_DeclinedTools.bas` - ✅ Tools for Admin to view/edit declined records (improved UX: positioned after Kartei, consistent headers/widths, last Decl_n comment only, AutoFilter, auto-cleanup)
- `Export_DeclinedHelpers.bas` - ✅ Helper functions for declined records management (comparison, copying, table creation)

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

```
[Ruck: ]Mnt.N: War(X); Ist(Y). [FieldName: Was(X); Is(Y). ]/Comment/ dd.mm.yyyy ||
```

**Supported fields:**

- `Mnt.1-12` → Monthly charges (columns U-AF)
- `Address` → Column F
- `Subject1` → Column J
- `Subject2` → Column O
- `Decl_N` → Decline comments from Superadmin

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
5. Marks each record in column X:
   - `Approved` → Move to tblKartei
   - `Declined` → Move to decl_tblKartei (with comment in column Y)
6. Runs `SyncDecisions` → Processes all decisions
   - In Mode A: If multiple blocks exist for same ID, only the decision from the last event is applied
   - In Mode B: Only one block per ID exists by design

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
