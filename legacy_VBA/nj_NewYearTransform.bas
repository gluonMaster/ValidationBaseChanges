Attribute VB_Name = "nj_NewYearTransform"
Option Explicit

' ============================================================================
' Module: nj_NewYearTransform
' Purpose: Core business logic for Kartei data transformation at year-end.
'          Prepares data for the new year by shifting fields, applying pricing
'          logic, and cleaning up obsolete records.
' ============================================================================

' Column indices constants for readability
Private Const COL_FAMILYID As Long = 1       ' A - FamilyID
Private Const COL_CHILD As Long = 4          ' D - Child name / "Zahlung"
Private Const COL_J As Long = 10             ' J - Subject1 (target)
Private Const COL_K As Long = 11             ' K
Private Const COL_L As Long = 12             ' L
Private Const COL_M As Long = 13             ' M - Price1 (target)
Private Const COL_N As Long = 14             ' N
Private Const COL_O As Long = 15             ' O - Subject2 (source for J)
Private Const COL_P As Long = 16             ' P
Private Const COL_Q As Long = 17             ' Q
Private Const COL_R As Long = 18             ' R - Price2 (source for M)
Private Const COL_S As Long = 19             ' S - Contains O/V marker
Private Const COL_T As Long = 20             ' T - KN marker column
Private Const COL_U As Long = 21             ' U - Month 1 (Jan)
Private Const COL_AA As Long = 27            ' AA - Month 7
Private Const COL_AB As Long = 28            ' AB - Month 8
Private Const COL_AE As Long = 31            ' AE - Month 11
Private Const COL_AF As Long = 32            ' AF - Month 12
Private Const COL_AK As Long = 37            ' AK - Extra1 (target)
Private Const COL_AL As Long = 38            ' AL
Private Const COL_AM As Long = 39            ' AM
Private Const COL_AN As Long = 40            ' AN
Private Const COL_AO As Long = 41            ' AO
Private Const COL_AP As Long = 42            ' AP - Extra source
Private Const COL_AQ As Long = 43            ' AQ
Private Const COL_AR As Long = 44            ' AR
Private Const COL_AS As Long = 45            ' AS
Private Const COL_AT As Long = 46            ' AT
Private Const COL_AV As Long = 48            ' AV - ID (never modify)
Private Const COL_BA As Long = 53            ' BA - Status (never modify)

Private Const PRICE_EPSILON As Double = 0.0001  ' Tolerance for price comparison

' ============================================================================
' PUBLIC ENTRY POINT
' ============================================================================

Public Sub nj_PrepareKarteiDataForNewYear(ByVal targetYear As Long)
    ' Main entry point for new year transformation.
    ' Operates directly on Kartei worksheet.
    ' targetYear parameter reserved for future use (logging, validation).
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = Worksheets("Kartei")
    
    ' Step 1: Delete rows with KN marker in column T
    nj_DeleteRowsWithKN ws
    
    ' Step 2: Delete orphan Zahlung rows (no content rows for FamilyID)
    nj_DeleteOrphanZahlungRows ws
    
    ' Step 3: Transform content rows for new year
    lastRow = indx_FindLastDataRow(ws)
    nj_TransformContentRows ws, lastRow
End Sub

' ============================================================================
' PRIVATE HELPER: DELETE ROWS WITH KN MARKER
' ============================================================================

Private Sub nj_DeleteRowsWithKN(ws As Worksheet)
    ' Deletes all rows where column T contains "KN" (case-insensitive).
    ' Iterates from bottom to top to avoid row index shifting issues.
    
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim cellValue As String
    
    lastRow = indx_FindLastDataRow(ws)
    
    For rowIndex = lastRow To 3 Step -1
        cellValue = CStr(ws.Cells(rowIndex, COL_T).Value)
        If indx_ContainsText(cellValue, "KN") Then
            ws.Rows(rowIndex).Delete
        End If
    Next rowIndex
End Sub

' ============================================================================
' PRIVATE HELPER: DELETE ORPHAN ZAHLUNG ROWS
' ============================================================================

Private Sub nj_DeleteOrphanZahlungRows(ws As Worksheet)
    ' Deletes Zahlung rows that have no corresponding content rows.
    ' A FamilyID is "orphan" if ALL its rows have D = "Zahlung".
    
    Dim familyHasContent As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim familyID As String
    
    Set familyHasContent = nj_BuildFamilyContentMap(ws)
    
    ' Second pass: delete orphan Zahlung rows (bottom to top)
    lastRow = indx_FindLastDataRow(ws)
    
    For rowIndex = lastRow To 3 Step -1
        If nj_IsZahlungRow(ws, rowIndex) Then
            familyID = nj_GetFamilyID(ws, rowIndex)
            If familyID <> "" Then
                If familyHasContent.Exists(familyID) Then
                    If familyHasContent(familyID) = False Then
                        ws.Rows(rowIndex).Delete
                    End If
                End If
            End If
        End If
    Next rowIndex
End Sub

Private Function nj_BuildFamilyContentMap(ws As Worksheet) As Object
    ' Builds a dictionary mapping FamilyID -> Boolean (has content rows).
    ' Returns Scripting.Dictionary with familyHasContent(FamilyID) = True/False.
    
    Dim dict As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim familyID As String
    Dim isZahlung As Boolean
    
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = indx_FindLastDataRow(ws)
    
    For rowIndex = 3 To lastRow
        familyID = nj_GetFamilyID(ws, rowIndex)
        If familyID <> "" Then
            isZahlung = nj_IsZahlungRow(ws, rowIndex)
            
            If Not dict.Exists(familyID) Then
                ' First occurrence: set based on whether this is content row
                dict.Add familyID, Not isZahlung
            Else
                ' If we find any content row, mark family as having content
                If Not isZahlung Then
                    dict(familyID) = True
                End If
            End If
        End If
    Next rowIndex
    
    Set nj_BuildFamilyContentMap = dict
End Function

Private Function nj_GetFamilyID(ws As Worksheet, rowIndex As Long) As String
    ' Extracts trimmed FamilyID from column A.
    nj_GetFamilyID = Trim$(CStr(ws.Cells(rowIndex, COL_FAMILYID).Value))
End Function

Private Function nj_IsZahlungRow(ws As Worksheet, rowIndex As Long) As Boolean
    ' Checks if row is a Zahlung row (column D = "Zahlung").
    ' Note: comparison includes leading space as per existing data format.
    nj_IsZahlungRow = (Trim$(CStr(ws.Cells(rowIndex, COL_CHILD).Value)) = "Zahlung")
End Function

' ============================================================================
' PRIVATE HELPER: TRANSFORM CONTENT ROWS
' ============================================================================

Private Sub nj_TransformContentRows(ws As Worksheet, lastRow As Long)
    ' Iterates through all content rows (D <> "Zahlung") and applies transformation.
    
    Dim rowIndex As Long
    
    For rowIndex = 3 To lastRow
        If Not nj_IsZahlungRow(ws, rowIndex) Then
            nj_TransformRowForNewYear ws, rowIndex
        End If
    Next rowIndex
End Sub

Private Sub nj_TransformRowForNewYear(ws As Worksheet, rowIndex As Long)
    ' Applies all new year transformations to a single content row.
    
    Dim oldO As Variant, oldP As Variant, oldQ As Variant, oldR As Variant, oldS As Variant
    
    ' 4.1: Read source values before overwriting
    oldO = ws.Cells(rowIndex, COL_O).Value
    oldP = ws.Cells(rowIndex, COL_P).Value
    oldQ = ws.Cells(rowIndex, COL_Q).Value
    oldR = ws.Cells(rowIndex, COL_R).Value
    oldS = ws.Cells(rowIndex, COL_S).Value
    
    ' 4.1: Shift J..N from O..S
    nj_ShiftFieldsJtoN ws, rowIndex, oldO, oldP, oldQ, oldR, oldS
    
    ' 4.3: Check special case for NH/Nachhilfe/Ind./VSpE in O (now read as oldO)
    If nj_IsNachhilfeRow(CStr(oldO)) Then
        nj_ZeroAllMonths ws, rowIndex
    Else
        ' 4.4: Apply month logic based on O/V marker
        nj_ApplyMonthLogic ws, rowIndex, oldR, oldS
    End If
    
    ' 4.5: Clear O..S and zero AB..AF
    nj_ClearSourceFields ws, rowIndex
    
    ' 4.6: Shift service columns AK..AO from AP..AT
    nj_ShiftServiceColumns ws, rowIndex
End Sub

' ============================================================================
' PRIVATE HELPER: FIELD SHIFTING
' ============================================================================

Private Sub nj_ShiftFieldsJtoN(ws As Worksheet, rowIndex As Long, _
                                oldO As Variant, oldP As Variant, oldQ As Variant, _
                                oldR As Variant, oldS As Variant)
    ' Shifts O..S values into J..N positions.
    
    ws.Cells(rowIndex, COL_J).Value = oldO
    ws.Cells(rowIndex, COL_K).Value = oldP
    ws.Cells(rowIndex, COL_L).Value = oldQ
    ws.Cells(rowIndex, COL_M).Value = oldR
    ws.Cells(rowIndex, COL_N).Value = oldS
End Sub

Private Sub nj_ShiftServiceColumns(ws As Worksheet, rowIndex As Long)
    ' Shifts AP..AT values into AK..AO, then clears AP..AT.
    
    ws.Cells(rowIndex, COL_AK).Value = ws.Cells(rowIndex, COL_AP).Value
    ws.Cells(rowIndex, COL_AL).Value = ws.Cells(rowIndex, COL_AQ).Value
    ws.Cells(rowIndex, COL_AM).Value = ws.Cells(rowIndex, COL_AR).Value
    ws.Cells(rowIndex, COL_AN).Value = ws.Cells(rowIndex, COL_AS).Value
    ws.Cells(rowIndex, COL_AO).Value = ws.Cells(rowIndex, COL_AT).Value
    
    ' Clear source columns
    ws.Cells(rowIndex, COL_AP).Value = ""
    ws.Cells(rowIndex, COL_AQ).Value = ""
    ws.Cells(rowIndex, COL_AR).Value = ""
    ws.Cells(rowIndex, COL_AS).Value = ""
    ws.Cells(rowIndex, COL_AT).Value = ""
End Sub

' ============================================================================
' PRIVATE HELPER: NACHHILFE DETECTION
' ============================================================================

Private Function nj_IsNachhilfeRow(subjectText As String) As Boolean
    ' Checks if subject indicates Nachhilfe/NH/Ind./VSpE special case.
    ' NH check is case-sensitive as per requirements.
    
    nj_IsNachhilfeRow = False
    
    ' Case-sensitive check for "NH"
    If InStr(subjectText, "NH") > 0 Then
        nj_IsNachhilfeRow = True
        Exit Function
    End If
    
    ' Case-insensitive checks for other markers
    If indx_ContainsText(subjectText, "Nachhilfe") Then
        nj_IsNachhilfeRow = True
        Exit Function
    End If
    
    If indx_ContainsText(subjectText, "Ind.") Then
        nj_IsNachhilfeRow = True
        Exit Function
    End If
    
    If indx_ContainsText(subjectText, "VSpE") Then
        nj_IsNachhilfeRow = True
        Exit Function
    End If
End Function

Private Sub nj_ZeroAllMonths(ws As Worksheet, rowIndex As Long)
    ' Sets all month columns U..AF (21..32) to zero.
    
    Dim colIndex As Long
    
    For colIndex = COL_U To COL_AF
        ws.Cells(rowIndex, colIndex).Value = 0
    Next colIndex
End Sub

' ============================================================================
' PRIVATE HELPER: MONTH LOGIC (O/V MARKER)
' ============================================================================

Private Sub nj_ApplyMonthLogic(ws As Worksheet, rowIndex As Long, _
                                oldR As Variant, oldS As Variant)
    ' Applies pricing logic to months U..AA based on O/V marker in S column.
    '
    ' Empty/invalid value handling:
    '   - Empty or non-numeric R -> priceR = 0
    '   - Empty or non-numeric AE -> priceAE = 0
    '   - Empty or non-numeric AF -> priceAF = 0
    '   - If all prices are 0 (or empty), months will be set to 0
    '   - "Equal" comparison uses PRICE_EPSILON tolerance (0.0001)
    '
    ' Note: This procedure never modifies columns AV (ID) or BA (Status).
    
    Dim hasOV As Boolean
    Dim priceR As Double
    Dim priceAE As Double
    Dim priceAF As Double
    
    hasOV = indx_ContainsText(CStr(oldS), "O/V")
    
    ' Get numeric values (empty/invalid cells return 0)
    priceR = nj_GetNumericPriceFromCell(ws.Cells(rowIndex, COL_R))
    priceAE = nj_SafeGetCellDouble(ws.Cells(rowIndex, COL_AE))
    priceAF = nj_SafeGetCellDouble(ws.Cells(rowIndex, COL_AF))
    
    If hasOV Then
        nj_ApplyOVBranchLogic ws, rowIndex, priceR, priceAE, priceAF
    Else
        nj_ApplyNonOVBranchLogic ws, rowIndex, priceR, priceAF
    End If
End Sub

Private Sub nj_ApplyOVBranchLogic(ws As Worksheet, rowIndex As Long, _
                                   priceR As Double, priceAE As Double, priceAF As Double)
    ' Branch A: O/V marker is present.
    ' U gets AF value, V..Z get priceR or priceAE, AA = 0.
    '
    ' Empty value behavior:
    '   - If R and AE are both 0 (empty), V..Z = 0, M unchanged
    '   - If R = 0 but AE > 0, V..Z = AE, M = AE (price changed)
    '   - If R > 0 but AE = 0, V..Z = AE (0), M = AE (price changed to 0)
    '   - If AF = 0 (empty), U = 0
    
    Dim colIndex As Long
    Dim usePrice As Double
    
    ' U (21) gets AF value (0 if AF was empty)
    ws.Cells(rowIndex, COL_U).Value = priceAF
    
    ' Determine price for V..Z using epsilon comparison
    ' If prices are equal (including both = 0), use priceR and don't update M
    If nj_PricesAreEqual(priceR, priceAE) Then
        usePrice = priceR
    Else
        usePrice = priceAE
        ' Update M with new base price (may be 0 if AE was empty)
        ws.Cells(rowIndex, COL_M).Value = priceAE
    End If
    
    ' V..Z (22..26)
    For colIndex = 22 To 26
        ws.Cells(rowIndex, colIndex).Value = usePrice
    Next colIndex
    
    ' AA (27) = 0
    ws.Cells(rowIndex, COL_AA).Value = 0
End Sub

Private Sub nj_ApplyNonOVBranchLogic(ws As Worksheet, rowIndex As Long, _
                                      priceR As Double, priceAF As Double)
    ' Branch B: O/V marker is NOT present.
    ' U..AA get priceR or priceAF.
    '
    ' Empty value behavior:
    '   - If R and AF are both 0 (empty), U..AA = 0, M unchanged
    '   - If R = 0 but AF > 0, U..AA = AF, M = AF (price changed)
    '   - If R > 0 but AF = 0, U..AA = AF (0), M = AF (price changed to 0)
    
    Dim colIndex As Long
    Dim usePrice As Double
    
    ' Determine price for U..AA using epsilon comparison
    ' If prices are equal (including both = 0), use priceR and don't update M
    If nj_PricesAreEqual(priceR, priceAF) Then
        usePrice = priceR
    Else
        usePrice = priceAF
        ' Update M with new base price (may be 0 if AF was empty)
        ws.Cells(rowIndex, COL_M).Value = priceAF
    End If
    
    ' U..AA (21..27)
    For colIndex = COL_U To COL_AA
        ws.Cells(rowIndex, colIndex).Value = usePrice
    Next colIndex
End Sub

' ============================================================================
' PRIVATE HELPER: CLEANUP
' ============================================================================

Private Sub nj_ClearSourceFields(ws As Worksheet, rowIndex As Long)
    ' Clears O..S (15..19) and sets AB..AF (28..32) to zero.
    
    Dim colIndex As Long
    
    ' Clear O..S
    For colIndex = COL_O To COL_S
        ws.Cells(rowIndex, colIndex).Value = ""
    Next colIndex
    
    ' Zero AB..AF
    For colIndex = COL_AB To COL_AF
        ws.Cells(rowIndex, colIndex).Value = 0
    Next colIndex
End Sub

' ============================================================================
' PRIVATE HELPER: NUMERIC CONVERSION
' ============================================================================

Private Function nj_GetNumericPriceFromCell(cell As Range) As Double
    ' Converts cell value to Double, handling text with various decimal separators.
    '
    ' Empty/invalid value handling:
    '   - Empty cell -> returns 0
    '   - Whitespace only -> returns 0
    '   - Non-numeric text -> returns 0
    '   - Numeric text with wrong decimal separator -> normalized and converted
    '
    ' This function is used for column R which often stores prices as text.
    ' Normalizes decimal separator based on system locale before conversion.
    
    Dim rawValue As String
    Dim systemDecimalSep As String
    Dim normalizedValue As String
    
    rawValue = Trim$(CStr(cell.Value))
    
    ' Empty or whitespace-only -> 0
    If Len(rawValue) = 0 Then
        nj_GetNumericPriceFromCell = 0
        Exit Function
    End If
    
    ' Get system decimal separator
    systemDecimalSep = Application.International(xlDecimalSeparator)
    
    ' Normalize decimal separator
    normalizedValue = rawValue
    If systemDecimalSep = "," Then
        normalizedValue = Replace(normalizedValue, ".", ",")
    Else
        normalizedValue = Replace(normalizedValue, ",", ".")
    End If
    
    ' Attempt conversion
    On Error Resume Next
    nj_GetNumericPriceFromCell = CDbl(normalizedValue)
    If Err.Number <> 0 Then
        nj_GetNumericPriceFromCell = 0
    End If
    On Error GoTo 0
End Function

Private Function nj_SafeGetCellDouble(cell As Range) As Double
    ' Safely retrieves cell value as Double from cells AE/AF.
    '
    ' Empty/invalid value handling:
    '   - Empty cell -> returns 0
    '   - Null value -> returns 0
    '   - Non-numeric value -> returns 0
    '   - Numeric value -> returns CDbl(value)
    '
    ' Assumes cell is already formatted as numeric by ConvertAndFormatCellsOptimized.
    ' Used for columns AE/AF which should be properly formatted after import.
    
    On Error Resume Next
    If IsEmpty(cell.Value) Or IsNull(cell.Value) Then
        nj_SafeGetCellDouble = 0
    ElseIf IsNumeric(cell.Value) Then
        nj_SafeGetCellDouble = CDbl(cell.Value)
    Else
        nj_SafeGetCellDouble = 0
    End If
    On Error GoTo 0
End Function

Private Function nj_PricesAreEqual(price1 As Double, price2 As Double) As Boolean
    ' Compares two prices with tolerance for floating point precision.
    '
    ' Uses PRICE_EPSILON (0.0001) to handle floating point rounding errors.
    ' This means:
    '   - 0 and 0 are equal (both empty/zero prices)
    '   - 10.50 and 10.50 are equal
    '   - 10.50 and 10.500001 are equal (within epsilon)
    '   - 10.50 and 10.51 are NOT equal
    '
    ' When both prices are 0 (e.g., R and AE both empty), this returns True,
    ' so the "equal" branch is taken and M is not updated.
    
    nj_PricesAreEqual = (Abs(price1 - price2) < PRICE_EPSILON)
End Function

' ============================================================================
' DEBUG HELPER: POST-TRANSFORMATION STATE CHECK
' ============================================================================

Public Sub nj_DebugCheckNewYearState()
    ' Debug helper to verify Kartei state after ImNeuenJahr_26 execution.
    ' Outputs results to Immediate Window (Debug.Print).
    ' Does NOT display any MsgBox dialogs.
    '
    ' Checks performed:
    '   1. No rows with "KN" marker in column T
    '   2. No orphan FamilyIDs (only Zahlung rows, no content rows)
    '   3. Numeric values in month columns U..AF for first 10 data rows
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Kartei")
    lastRow = indx_FindLastDataRow(ws)
    
    Debug.Print "=========================================="
    Debug.Print "nj_DebugCheckNewYearState - " & Now()
    Debug.Print "=========================================="
    Debug.Print "Total data rows: " & (lastRow - 2)
    Debug.Print ""
    
    ' Check 1: No KN markers
    nj_DebugCheckNoKNMarkers ws, lastRow
    
    ' Check 2: No orphan Zahlung families
    nj_DebugCheckNoOrphanZahlung ws, lastRow
    
    ' Check 3: Numeric values in month columns for first 10 rows
    nj_DebugCheckMonthsNumeric ws, lastRow
    
    Debug.Print "=========================================="
    Debug.Print "Debug check completed."
    Debug.Print "=========================================="
End Sub

Private Sub nj_DebugCheckNoKNMarkers(ws As Worksheet, lastRow As Long)
    ' Checks that no rows contain "KN" in column T.
    
    Dim rowIndex As Long
    Dim cellValue As String
    Dim knCount As Long
    
    knCount = 0
    
    For rowIndex = 3 To lastRow
        cellValue = CStr(ws.Cells(rowIndex, COL_T).Value)
        If indx_ContainsText(cellValue, "KN") Then
            knCount = knCount + 1
            Debug.Print "  [FAIL] KN marker found in row " & rowIndex & ": " & cellValue
        End If
    Next rowIndex
    
    If knCount = 0 Then
        Debug.Print "[PASS] Check 1: No KN markers found in column T"
    Else
        Debug.Print "[FAIL] Check 1: Found " & knCount & " rows with KN marker"
    End If
    Debug.Print ""
End Sub

Private Sub nj_DebugCheckNoOrphanZahlung(ws As Worksheet, lastRow As Long)
    ' Checks that no FamilyID has only Zahlung rows without content rows.
    
    Dim dict As Object
    Dim rowIndex As Long
    Dim familyID As String
    Dim isZahlung As Boolean
    Dim orphanCount As Long
    Dim key As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Build family content map
    For rowIndex = 3 To lastRow
        familyID = Trim$(CStr(ws.Cells(rowIndex, COL_FAMILYID).Value))
        If familyID <> "" Then
            isZahlung = (Trim$(CStr(ws.Cells(rowIndex, COL_CHILD).Value)) = "Zahlung")
            
            If Not dict.Exists(familyID) Then
                dict.Add familyID, Not isZahlung
            Else
                If Not isZahlung Then
                    dict(familyID) = True
                End If
            End If
        End If
    Next rowIndex
    
    ' Check for orphan families
    orphanCount = 0
    For Each key In dict.Keys
        If dict(key) = False Then
            orphanCount = orphanCount + 1
            Debug.Print "  [FAIL] Orphan FamilyID (only Zahlung): " & key
        End If
    Next key
    
    If orphanCount = 0 Then
        Debug.Print "[PASS] Check 2: No orphan Zahlung families found"
    Else
        Debug.Print "[FAIL] Check 2: Found " & orphanCount & " orphan families"
    End If
    Debug.Print ""
End Sub

Private Sub nj_DebugCheckMonthsNumeric(ws As Worksheet, lastRow As Long)
    ' Checks that month columns U..AF contain numeric values for first 10 data rows.
    
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim checkRows As Long
    Dim cellValue As Variant
    Dim nonNumericCount As Long
    Dim rowsChecked As Long
    
    ' Check up to 10 rows (or less if fewer data rows)
    checkRows = Application.WorksheetFunction.Min(10, lastRow - 2)
    nonNumericCount = 0
    rowsChecked = 0
    
    For rowIndex = 3 To 3 + checkRows - 1
        ' Skip Zahlung rows for this check
        If Trim$(CStr(ws.Cells(rowIndex, COL_CHILD).Value)) <> "Zahlung" Then
            rowsChecked = rowsChecked + 1
            
            For colIndex = COL_U To COL_AF
                cellValue = ws.Cells(rowIndex, colIndex).Value
                
                ' Empty cells are acceptable (will be treated as 0)
                If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
                    If Not IsNumeric(cellValue) Then
                        nonNumericCount = nonNumericCount + 1
                        Debug.Print "  [FAIL] Non-numeric value in row " & rowIndex & _
                                    ", col " & colIndex & ": " & cellValue
                    End If
                End If
            Next colIndex
        End If
    Next rowIndex
    
    If nonNumericCount = 0 Then
        Debug.Print "[PASS] Check 3: Month columns U..AF are numeric (" & rowsChecked & " rows checked)"
    Else
        Debug.Print "[FAIL] Check 3: Found " & nonNumericCount & " non-numeric values in months"
    End If
    Debug.Print ""
End Sub
