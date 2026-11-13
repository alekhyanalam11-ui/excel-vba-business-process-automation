Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 1. TRIGGER CHECK
    If Target.Address <> "$Q$1" Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim lastRow As Long, r As Long
    Dim tbl As ListObject
    Dim monthKey As String
    Dim monthDict As Object
    Set monthDict = CreateObject("Scripting.Dictionary")
    
    ' 2. RESET TABLE: Remove Totals to identify real data rows
    If Me.ListObjects.Count > 0 Then
        Set tbl = Me.ListObjects(1)
        tbl.ShowTotals = False
    End If

    lastRow = Me.Cells(Me.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then GoTo SafeExit

    ' 3. MONTHLY & ZERO-DEBIT LOGIC (Column G is P DEBIT)
    
    ' First Pass: Identify months that have a total P DEBIT > 0
    For r = 4 To lastRow
        If IsDate(Me.Cells(r, "A").Value) Then
            ' Create a unique key for Month-Year
            monthKey = Month(Me.Cells(r, "A").Value) & "-" & Year(Me.Cells(r, "A").Value)
            ' Sum the values in Column G (7)
            monthDict(monthKey) = monthDict(monthKey) + Val(Me.Cells(r, "G").Value)
        End If
    Next r

    ' Second Pass: Delete rows based on the logic
    For r = lastRow To 4 Step -1
        monthKey = ""
        If IsDate(Me.Cells(r, "A").Value) Then
            monthKey = Month(Me.Cells(r, "A").Value) & "-" & Year(Me.Cells(r, "A").Value)
        End If

        '-------------------------------------------------
        ' Delete Logic
        '-------------------------------------------------

If monthKey <> "" Then

    'Delete the month ONLY if monthly balance is exactly ZERO
    If Abs(monthDict(monthKey)) < 0.01 Then

        Me.Rows(r).Delete

    Else

        'Single Entry:
        'Delete ONLY rows whose P Debit is exactly ZERO.
        'Keep both Positive and Negative values.

        If IsNumeric(Me.Cells(r, "G").Value) Then

            If CDbl(Me.Cells(r, "G").Value) = 0 Then

                Me.Rows(r).Delete

            End If

        End If

    End If

End If
    Next r

    ' Re-calculate last row after deletion
    lastRow = Me.Cells(Me.Rows.Count, "A").End(xlUp).Row

    ' 4. APPLY FORMATTING & RECALCULATE TOTALS
    ' Check if there is any data left (excluding header)
    If lastRow < 4 Or (lastRow = 4 And Me.Cells(4, 1).Value = "") Then
        If lastRow >= 4 Then Me.Rows("4:" & lastRow).Delete
        Me.Range("A4").Value = "NO DATA"
        Me.Range("A4").Font.Bold = True
        MsgBox "NO DATA FOUND", vbInformation
    Else
        ' Force Worksheet Bold for columns A through G
        Me.Range("A1:G" & lastRow).Font.Bold = True
        
        If Not tbl Is Nothing Then
            ' Resize table to the remaining data so Totals reflect only visible rows
            tbl.Resize Me.Range("A3:G" & lastRow)
            tbl.Range.Font.Bold = True
            tbl.ShowTotals = True
            
            ' Re-apply calculations for P-Debit and other columns
            On Error Resume Next
            tbl.ListColumns("DELIVERY QTY").TotalsCalculation = xlTotalsCalculationSum
            tbl.ListColumns(5).TotalsCalculation = xlTotalsCalculationSum
            tbl.ListColumns(6).TotalsCalculation = xlTotalsCalculationSum
            tbl.ListColumns(7).TotalsCalculation = xlTotalsCalculationSum ' P DEBIT
            
            If Not tbl.TotalsRowRange Is Nothing Then tbl.TotalsRowRange.Font.Bold = True
            On Error GoTo SafeExit
        End If
        MsgBox "NON-LOCAL FILTERED & BOLDED", vbInformation
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
