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
    
    ' 2. RESET TABLE
    If Me.ListObjects.Count > 0 Then
        Set tbl = Me.ListObjects(1)
        tbl.ShowTotals = False
    End If

    lastRow = Me.Cells(Me.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then GoTo SafeExit

    ' 3. MONTHLY & ZERO-DEBIT LOGIC
    ' First Pass: Identify months that have a total due > 0
    For r = 4 To lastRow
        If IsDate(Me.Cells(r, "A").Value) Then
            monthKey = Month(Me.Cells(r, "A").Value) & "-" & Year(Me.Cells(r, "A").Value)
            monthDict(monthKey) = monthDict(monthKey) + Val(Me.Cells(r, "J").Value)
        End If
    Next r

    ' Second Pass: Delete rows
    For r = lastRow To 4 Step -1
        monthKey = ""
        If IsDate(Me.Cells(r, "A").Value) Then
            monthKey = Month(Me.Cells(r, "A").Value) & "-" & Year(Me.Cells(r, "A").Value)
        End If

        ' DELETE IF:
        ' 1. The entire month has 0 total debit OR
        ' 2. This specific row has 0 debit (as per your original requirement)
        If monthKey <> "" Then

    If Abs(monthDict(monthKey)) < 0.01 Then

        Me.Rows(r).Delete

    ElseIf IsNumeric(Me.Cells(r, "J").Value) Then

        If CDbl(Me.Cells(r, "J").Value) = 0 Then
            Me.Rows(r).Delete
        End If

    End If

End If
    Next r

    lastRow = Me.Cells(Me.Rows.Count, "A").End(xlUp).Row

    ' 4. APPLY FORMATTING & RECALCULATE TOTALS
    If lastRow < 4 Or (lastRow = 4 And Me.Cells(4, 1).Value = "") Then
        If lastRow >= 4 Then Me.Rows("4:" & lastRow).Delete
        Me.Range("A4").Value = "NO DATA"
        Me.Range("A4").Font.Bold = True
        MsgBox "NO DATA FOUND", vbInformation
    Else
        Me.Range("A1:J" & lastRow).Font.Bold = True
        
        If Not tbl Is Nothing Then
            ' Resize table to the remaining data so Totals reflect only visible rows
            tbl.Resize Me.Range("A3:J" & lastRow)
            tbl.Range.Font.Bold = True
            tbl.ShowTotals = True
            
            On Error Resume Next
            tbl.ListColumns("DELIVERY QTY").TotalsCalculation = xlTotalsCalculationSum
            tbl.ListColumns("TOTAL").TotalsCalculation = xlTotalsCalculationSum
            tbl.ListColumns("CREDIT").TotalsCalculation = xlTotalsCalculationSum
            tbl.ListColumns("DEBIT").TotalsCalculation = xlTotalsCalculationSum
            If Not tbl.TotalsRowRange Is Nothing Then tbl.TotalsRowRange.Font.Bold = True
            On Error GoTo SafeExit
        End If
        MsgBox "LOCAL FILTERED & BOLDED", vbInformation
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
