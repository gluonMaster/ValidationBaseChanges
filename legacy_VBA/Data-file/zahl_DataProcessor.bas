Attribute VB_Name = "zahl_DataProcessor"
' Data processing module for family payments
' Handles the logic for categorizing and processing payment records

Option Explicit

Public Sub ProcessAllData()
    ' Main data processing procedure
    Dim i As Integer, j As Integer
    Dim cellValueA As Variant, cellValueB As Variant
    
    ' Initialize first family with safe value retrieval
    cellValueA = Worksheets("Kartei").Cells(3, "A").Value
    cellValueB = Worksheets("Kartei").Cells(3, "B").Value
    
    zahl_Main.num(1) = IIf(IsNull(cellValueA), "", CStr(cellValueA))
    zahl_Main.nam(1) = IIf(IsNull(cellValueB), "", CStr(cellValueB))
    j = 1
    
    ' Process all rows from Kartei worksheet
    For i = 4 To zahl_Main.letzRow
        cellValueA = Worksheets("Kartei").Cells(i, "A").Value
        Dim currentNum As String
        currentNum = IIf(IsNull(cellValueA), "", CStr(cellValueA))
        
        If zahl_Main.num(j) = currentNum Then
            ' Same family - process the payment record
            ' Check for name consistency
            cellValueB = Worksheets("Kartei").Cells(i, "B").Value
            Dim currentName As String
            currentName = IIf(IsNull(cellValueB), "", CStr(cellValueB))
            
            If zahl_Main.nam(j) <> currentName And Len(Trim(currentName)) > 0 Then
                ' Name inconsistency detected - will be caught by validation
                ' For now, keep the first name encountered
            End If
            
            Call ProcessSingleRecord(i, j)
        Else
            ' New family found
            j = j + 1
            If j > 2000 Then
                MsgBox "Zu viele Familien (mehr als 2000). Verarbeitung gestoppt.", vbCritical, "Fehler"
                Exit Sub
            End If
            
            cellValueB = Worksheets("Kartei").Cells(i, "B").Value
            zahl_Main.num(j) = currentNum
            zahl_Main.nam(j) = IIf(IsNull(cellValueB), "", CStr(cellValueB))
            Call ProcessSingleRecord(i, j)
        End If
    Next i
End Sub

Private Sub ProcessSingleRecord(row As Integer, familyIndex As Integer)
    ' Process a single payment record and categorize it
    Dim colJ As String, colO As String
    Dim hasNachhilfeJ As Boolean, hasNachhilfeO As Boolean
    Dim hasIndividualJ As Boolean, hasIndividualO As Boolean
    Dim recordType As Integer
    Dim cellValueJ As Variant, cellValueO As Variant
    
    ' Get values from columns J and O with safe conversion
    cellValueJ = Worksheets("Kartei").Cells(row, "J").Value
    cellValueO = Worksheets("Kartei").Cells(row, "O").Value
    
    ' Convert to string safely
    If IsNull(cellValueJ) Or IsEmpty(cellValueJ) Then
        colJ = ""
    Else
        colJ = CStr(cellValueJ)
    End If
    
    If IsNull(cellValueO) Or IsEmpty(cellValueO) Then
        colO = ""
    Else
        colO = CStr(cellValueO)
    End If
    
    ' Check for markers in both columns
    hasNachhilfeJ = HasNachhilfeMarker(colJ)
    hasNachhilfeO = HasNachhilfeMarker(colO)
    hasIndividualJ = HasIndividualMarker(colJ)
    hasIndividualO = HasIndividualMarker(colO)
    
    ' Determine record type and process accordingly
    recordType = DetermineRecordType(hasNachhilfeJ, hasNachhilfeO, hasIndividualJ, hasIndividualO)
    
    Select Case recordType
        Case 0 ' Regular costs only
            Call AddToRegularCosts(row, familyIndex, False)
            
        Case 1 ' Pure Nachhilfe (both semesters)
            Call AddToNachhilfeCosts(row, familyIndex, False)
            
        Case 2 ' Pure Individual (both semesters)
            Call AddToIndividualCosts(row, familyIndex, False)
            
        Case 3 ' Mixed Nachhilfe (only first semester)
            Call AddToNachhilfeCosts(row, familyIndex, True, True, False) ' First semester to Nachhilfe
            Call AddToRegularCosts(row, familyIndex, True, False, True)   ' Second semester to Regular
            
        Case 4 ' Mixed Nachhilfe (only second semester)
            Call AddToRegularCosts(row, familyIndex, True, True, False)   ' First semester to Regular
            Call AddToNachhilfeCosts(row, familyIndex, True, False, True) ' Second semester to Nachhilfe
            
        Case 5 ' Mixed Individual (only first semester)
            Call AddToIndividualCosts(row, familyIndex, True, True, False) ' First semester to Individual
            Call AddToRegularCosts(row, familyIndex, True, False, True)    ' Second semester to Regular
            
        Case 6 ' Mixed Individual (only second semester)
            Call AddToRegularCosts(row, familyIndex, True, True, False)    ' First semester to Regular
            Call AddToIndividualCosts(row, familyIndex, True, False, True) ' Second semester to Individual
            
        Case 7 ' Conflict: Both Nachhilfe and Individual markers
            Call AddToIndividualCosts(row, familyIndex, False, False, False, True) ' Add to Individual with yellow marking
            zahl_Main.conflictCount = zahl_Main.conflictCount + 1
            
    End Select
End Sub

Private Function DetermineRecordType(nachhilfeJ As Boolean, nachhilfeO As Boolean, individualJ As Boolean, individualO As Boolean) As Integer
    ' Returns:
    ' 0 = Regular costs
    ' 1 = Pure Nachhilfe (both semesters)
    ' 2 = Pure Individual (both semesters)
    ' 3 = Mixed Nachhilfe (first semester only)
    ' 4 = Mixed Nachhilfe (second semester only)
    ' 5 = Mixed Individual (first semester only)
    ' 6 = Mixed Individual (second semester only)
    ' 7 = Conflict (both types present)
    
    ' Check for conflicts first
    If (nachhilfeJ Or nachhilfeO) And (individualJ Or individualO) Then
        DetermineRecordType = 7
        Exit Function
    End If
    
    ' Check for Individual lessons
    If individualJ And individualO Then
        DetermineRecordType = 2 ' Pure Individual
    ElseIf individualJ And Not individualO Then
        DetermineRecordType = 5 ' Mixed Individual (first semester)
    ElseIf Not individualJ And individualO Then
        DetermineRecordType = 6 ' Mixed Individual (second semester)
    ' Check for Nachhilfe
    ElseIf nachhilfeJ And nachhilfeO Then
        DetermineRecordType = 1 ' Pure Nachhilfe
    ElseIf nachhilfeJ And Not nachhilfeO Then
        DetermineRecordType = 3 ' Mixed Nachhilfe (first semester)
    ElseIf Not nachhilfeJ And nachhilfeO Then
        DetermineRecordType = 4 ' Mixed Nachhilfe (second semester)
    Else
        DetermineRecordType = 0 ' Regular costs
    End If
End Function

Private Function HasNachhilfeMarker(text As String) As Boolean
    ' Check for Nachhilfe markers (existing logic)
    HasNachhilfeMarker = (InStr(text, "Nach") > 0) Or (InStr(text, "NH") > 0)
End Function

Private Function HasIndividualMarker(text As String) As Boolean
    ' Check for Individual lesson markers with smart detection
    Dim upperText As String
    upperText = UCase(text)
    
    ' Check each marker with following space or punctuation
    If CheckMarkerWithSeparator(upperText, "IND ") Or _
       CheckMarkerWithSeparator(upperText, "IND.") Or _
       CheckMarkerWithSeparator(upperText, "VSPE") Or _
       InStr(upperText, "PSCHICHOLOGISCHE BERATUNG") > 0 Then
        HasIndividualMarker = True
    Else
        HasIndividualMarker = False
    End If
End Function

Private Function CheckMarkerWithSeparator(text As String, marker As String) As Boolean
    ' Check if marker exists followed by space, punctuation, or end of string
    Dim pos As Integer
    pos = InStr(text, marker)
    
    If pos > 0 Then
        Dim nextPos As Integer
        nextPos = pos + Len(marker)
        
        If nextPos > Len(text) Then
            ' Marker is at the end of string
            CheckMarkerWithSeparator = True
        Else
            Dim nextChar As String
            nextChar = Mid(text, nextPos, 1)
            ' Check if next character is space or punctuation
            If nextChar = " " Or nextChar = "." Or nextChar = "_" Or nextChar = "-" Or nextChar = "," Then
                CheckMarkerWithSeparator = True
            Else
                CheckMarkerWithSeparator = False
            End If
        End If
    Else
        CheckMarkerWithSeparator = False
    End If
End Function

Private Sub AddToRegularCosts(row As Integer, familyIndex As Integer, Optional mixed As Boolean = False, _
                             Optional firstSemester As Boolean = False, Optional secondSemester As Boolean = False)
    ' Add amounts to regular costs arrays
    Call AddAmountsToArrays(row, familyIndex, "regular", mixed, firstSemester, secondSemester)
End Sub

Private Sub AddToNachhilfeCosts(row As Integer, familyIndex As Integer, Optional mixed As Boolean = False, _
                               Optional firstSemester As Boolean = False, Optional secondSemester As Boolean = False)
    ' Add amounts to Nachhilfe costs arrays
    Call AddAmountsToArrays(row, familyIndex, "nachhilfe", mixed, firstSemester, secondSemester)
End Sub

Private Sub AddToIndividualCosts(row As Integer, familyIndex As Integer, Optional mixed As Boolean = False, _
                                Optional firstSemester As Boolean = False, Optional secondSemester As Boolean = False, _
                                Optional conflict As Boolean = False)
    ' Add amounts to Individual costs arrays
    Call AddAmountsToArrays(row, familyIndex, "individual", mixed, firstSemester, secondSemester, conflict)
End Sub

Private Sub AddAmountsToArrays(row As Integer, familyIndex As Integer, arrayType As String, _
                              Optional mixed As Boolean = False, Optional firstSemester As Boolean = False, _
                              Optional secondSemester As Boolean = False, Optional conflict As Boolean = False)
    ' Generic procedure to add amounts to specified arrays with color marking
    Dim amounts(1 To 12) As Double
    Dim i As Integer, colorValue As Integer
    Dim cellValue As Variant
    
    ' Get amounts from Excel row (columns 21-32) with safe conversion
    For i = 1 To 12
        cellValue = Worksheets("Kartei").Cells(row, 20 + i).Value
        If IsNumeric(cellValue) Then
            amounts(i) = CDbl(cellValue)
        Else
            amounts(i) = 0
        End If
    Next i
    
    ' Determine color value
    If conflict Then
        colorValue = 2 ' Yellow for conflicts
    Else
        colorValue = 0 ' Normal (removed pink for mixed cases)
    End If
    
    ' Add amounts to appropriate arrays
    Select Case LCase(arrayType)
        Case "regular"
            Call AddToSpecificArrays(amounts, familyIndex, "regular", mixed, firstSemester, secondSemester, colorValue)
            
        Case "nachhilfe"
            Call AddToSpecificArrays(amounts, familyIndex, "nachhilfe", mixed, firstSemester, secondSemester, colorValue)
            
        Case "individual"
            Call AddToSpecificArrays(amounts, familyIndex, "individual", mixed, firstSemester, secondSemester, colorValue)
    End Select
End Sub

Private Sub AddToSpecificArrays(amounts() As Double, familyIndex As Integer, arrayType As String, _
                               mixed As Boolean, firstSemester As Boolean, secondSemester As Boolean, colorValue As Integer)
    ' Add amounts to specific array type with proper handling of mixed cases
    Dim i As Integer
    
    Select Case LCase(arrayType)
        Case "regular"
            For i = 1 To 12
                If Not mixed Then
                    ' Normal case - add all months
                    Call AddToRegularMonth(familyIndex, i, amounts(i), colorValue)
                ElseIf (firstSemester And i <= 6) Or (secondSemester And i > 6) Then
                    ' Mixed case - add only relevant semester
                    Call AddToRegularMonth(familyIndex, i, amounts(i), colorValue)
                ' Removed pink marking for mixed cases
                End If
            Next i
            
        Case "nachhilfe"
            For i = 1 To 12
                If Not mixed Then
                    Call AddToNachhilfeMonth(familyIndex, i, amounts(i), colorValue)
                ElseIf (firstSemester And i <= 6) Or (secondSemester And i > 6) Then
                    Call AddToNachhilfeMonth(familyIndex, i, amounts(i), colorValue)
                ' Removed pink marking for mixed cases
                End If
            Next i
            
        Case "individual"
            For i = 1 To 12
                If Not mixed Then
                    Call AddToIndividualMonth(familyIndex, i, amounts(i), colorValue)
                ElseIf (firstSemester And i <= 6) Or (secondSemester And i > 6) Then
                    Call AddToIndividualMonth(familyIndex, i, amounts(i), colorValue)
                Else
                    ' Keep yellow marking for conflicts only
                    If colorValue = 2 Then zahl_Main.colorIndividual(familyIndex, i) = 2
                End If
            Next i
    End Select
End Sub

Private Sub AddToRegularMonth(familyIndex As Integer, month As Integer, amount As Double, colorValue As Integer)
    Select Case month
        Case 1: zahl_Main.sum01(familyIndex) = zahl_Main.sum01(familyIndex) + amount
        Case 2: zahl_Main.sum02(familyIndex) = zahl_Main.sum02(familyIndex) + amount
        Case 3: zahl_Main.sum03(familyIndex) = zahl_Main.sum03(familyIndex) + amount
        Case 4: zahl_Main.sum04(familyIndex) = zahl_Main.sum04(familyIndex) + amount
        Case 5: zahl_Main.sum05(familyIndex) = zahl_Main.sum05(familyIndex) + amount
        Case 6: zahl_Main.sum06(familyIndex) = zahl_Main.sum06(familyIndex) + amount
        Case 7: zahl_Main.sum07(familyIndex) = zahl_Main.sum07(familyIndex) + amount
        Case 8: zahl_Main.sum08(familyIndex) = zahl_Main.sum08(familyIndex) + amount
        Case 9: zahl_Main.sum09(familyIndex) = zahl_Main.sum09(familyIndex) + amount
        Case 10: zahl_Main.sum10(familyIndex) = zahl_Main.sum10(familyIndex) + amount
        Case 11: zahl_Main.sum11(familyIndex) = zahl_Main.sum11(familyIndex) + amount
        Case 12: zahl_Main.sum12(familyIndex) = zahl_Main.sum12(familyIndex) + amount
    End Select
    
    If colorValue > 0 Then zahl_Main.colorRegular(familyIndex, month) = colorValue
    If HasNonZeroValues(familyIndex, "regular") Then zahl_Main.nichtNull(familyIndex) = True
End Sub

Private Sub AddToNachhilfeMonth(familyIndex As Integer, month As Integer, amount As Double, colorValue As Integer)
    Select Case month
        Case 1: zahl_Main.sum01N(familyIndex) = zahl_Main.sum01N(familyIndex) + amount
        Case 2: zahl_Main.sum02N(familyIndex) = zahl_Main.sum02N(familyIndex) + amount
        Case 3: zahl_Main.sum03N(familyIndex) = zahl_Main.sum03N(familyIndex) + amount
        Case 4: zahl_Main.sum04N(familyIndex) = zahl_Main.sum04N(familyIndex) + amount
        Case 5: zahl_Main.sum05N(familyIndex) = zahl_Main.sum05N(familyIndex) + amount
        Case 6: zahl_Main.sum06N(familyIndex) = zahl_Main.sum06N(familyIndex) + amount
        Case 7: zahl_Main.sum07N(familyIndex) = zahl_Main.sum07N(familyIndex) + amount
        Case 8: zahl_Main.sum08N(familyIndex) = zahl_Main.sum08N(familyIndex) + amount
        Case 9: zahl_Main.sum09N(familyIndex) = zahl_Main.sum09N(familyIndex) + amount
        Case 10: zahl_Main.sum10N(familyIndex) = zahl_Main.sum10N(familyIndex) + amount
        Case 11: zahl_Main.sum11N(familyIndex) = zahl_Main.sum11N(familyIndex) + amount
        Case 12: zahl_Main.sum12N(familyIndex) = zahl_Main.sum12N(familyIndex) + amount
    End Select
    
    If colorValue > 0 Then zahl_Main.colorNachhilfe(familyIndex, month) = colorValue
    If HasNonZeroValues(familyIndex, "nachhilfe") Then zahl_Main.nichtNullN(familyIndex) = True
End Sub

Private Sub AddToIndividualMonth(familyIndex As Integer, month As Integer, amount As Double, colorValue As Integer)
    Select Case month
        Case 1: zahl_Main.sum01I(familyIndex) = zahl_Main.sum01I(familyIndex) + amount
        Case 2: zahl_Main.sum02I(familyIndex) = zahl_Main.sum02I(familyIndex) + amount
        Case 3: zahl_Main.sum03I(familyIndex) = zahl_Main.sum03I(familyIndex) + amount
        Case 4: zahl_Main.sum04I(familyIndex) = zahl_Main.sum04I(familyIndex) + amount
        Case 5: zahl_Main.sum05I(familyIndex) = zahl_Main.sum05I(familyIndex) + amount
        Case 6: zahl_Main.sum06I(familyIndex) = zahl_Main.sum06I(familyIndex) + amount
        Case 7: zahl_Main.sum07I(familyIndex) = zahl_Main.sum07I(familyIndex) + amount
        Case 8: zahl_Main.sum08I(familyIndex) = zahl_Main.sum08I(familyIndex) + amount
        Case 9: zahl_Main.sum09I(familyIndex) = zahl_Main.sum09I(familyIndex) + amount
        Case 10: zahl_Main.sum10I(familyIndex) = zahl_Main.sum10I(familyIndex) + amount
        Case 11: zahl_Main.sum11I(familyIndex) = zahl_Main.sum11I(familyIndex) + amount
        Case 12: zahl_Main.sum12I(familyIndex) = zahl_Main.sum12I(familyIndex) + amount
    End Select
    
    If colorValue > 0 Then zahl_Main.colorIndividual(familyIndex, month) = colorValue
    If HasNonZeroValues(familyIndex, "individual") Then zahl_Main.nichtNullI(familyIndex) = True
End Sub

Private Function HasNonZeroValues(familyIndex As Integer, arrayType As String) As Boolean
    ' Check if family has any non-zero values for specified array type
    Dim total As Double
    
    Select Case LCase(arrayType)
        Case "regular"
            total = zahl_Main.sum01(familyIndex) + zahl_Main.sum02(familyIndex) + zahl_Main.sum03(familyIndex) + _
                   zahl_Main.sum04(familyIndex) + zahl_Main.sum05(familyIndex) + zahl_Main.sum06(familyIndex) + _
                   zahl_Main.sum07(familyIndex) + zahl_Main.sum08(familyIndex) + zahl_Main.sum09(familyIndex) + _
                   zahl_Main.sum10(familyIndex) + zahl_Main.sum11(familyIndex) + zahl_Main.sum12(familyIndex)
                   
        Case "nachhilfe"
            total = zahl_Main.sum01N(familyIndex) + zahl_Main.sum02N(familyIndex) + zahl_Main.sum03N(familyIndex) + _
                   zahl_Main.sum04N(familyIndex) + zahl_Main.sum05N(familyIndex) + zahl_Main.sum06N(familyIndex) + _
                   zahl_Main.sum07N(familyIndex) + zahl_Main.sum08N(familyIndex) + zahl_Main.sum09N(familyIndex) + _
                   zahl_Main.sum10N(familyIndex) + zahl_Main.sum11N(familyIndex) + zahl_Main.sum12N(familyIndex)
                   
        Case "individual"
            total = zahl_Main.sum01I(familyIndex) + zahl_Main.sum02I(familyIndex) + zahl_Main.sum03I(familyIndex) + _
                   zahl_Main.sum04I(familyIndex) + zahl_Main.sum05I(familyIndex) + zahl_Main.sum06I(familyIndex) + _
                   zahl_Main.sum07I(familyIndex) + zahl_Main.sum08I(familyIndex) + zahl_Main.sum09I(familyIndex) + _
                   zahl_Main.sum10I(familyIndex) + zahl_Main.sum11I(familyIndex) + zahl_Main.sum12I(familyIndex)
    End Select
    
    HasNonZeroValues = (total > 0)
End Function
