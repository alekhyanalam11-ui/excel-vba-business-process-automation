Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim wsNonLocal As Worksheet
    Dim wsRes As Worksheet
    Dim r As Long, i As Long, j As Long
    Dim lastRow As Long, lastRowM As Long, lastRowN As Long
    Dim outRow As Long
    Dim cutoffDate As Date
    Dim millerName As String
    Dim tbl As ListObject
    Dim dataFound As Boolean
    Dim totalDue As Double

'========================
' K1 CLICK (UNCHANGED)
'========================
If Target.Address = "$K$1" Then

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Set wsNonLocal = Sheets("NONLOCAL")
    cutoffDate = Date - 1
    
    lastRowM = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
    lastRowN = wsNonLocal.Cells(wsNonLocal.Rows.Count, "C").End(xlUp).Row

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim key As String
    Dim amt As Double

    For j = 2 To lastRowN
        If wsNonLocal.Cells(j, "B").Value <= cutoffDate Then
            
            key = Trim(wsNonLocal.Cells(j, "C").Value)
            amt = Val(wsNonLocal.Cells(j, "T").Value)
            
            If key <> "" Then
                If dict.Exists(key) Then
                    dict(key) = dict(key) + amt
                Else
                    dict.Add key, amt
                End If
            End If
            
        End If
    Next j

    For i = 2 To lastRowM
        millerName = Trim(Me.Cells(i, "B").Value)
        
        If millerName <> "" Then
            If dict.Exists(millerName) Then
                Me.Cells(i, "D").Value = dict(millerName)
            Else
                Me.Cells(i, "D").Value = 0
            End If
        End If
    Next i

    MsgBox "Miller Total Due (Up to Yesterday) Updated Successfully"
    GoTo SafeExit

End If

'========================
' REPORT GENERATION (MONTH LOGIC ADDED)
'========================
If Target.Column <> 2 Then Exit Sub
If Target.Row = 1 Then Exit Sub

On Error GoTo SafeExit

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Set wsNonLocal = Sheets("NONLOCAL")
Set wsRes = Sheets("Results_NONLOCAL")

millerName = Target.Value
cutoffDate = Date - 1

wsRes.Range("A1:G1000").Clear

For Each tbl In wsRes.ListObjects
    tbl.Delete
Next tbl

With wsRes
    .Range("A1").Value = millerName
    .Range("A1:E1").Merge
    .Range("A1").HorizontalAlignment = xlCenter
    .Range("A1").Font.Size = 18
    .Range("A1").Font.Bold = True

    .Range("A3").Resize(1, 7).Value = _
    Array("DL DATE", "PARTY", "DELIVERY QTY", "CR MILLER", "M COM", "M CREDIT", "M DEBIT")

    .Rows(3).Font.Bold = True
End With

outRow = 4
lastRow = wsNonLocal.Cells(wsNonLocal.Rows.Count, "B").End(xlUp).Row

'========================
' ? MONTH FILTER (COLUMN T)
'========================
Dim dictMonth As Object
Set dictMonth = CreateObject("Scripting.Dictionary")

Dim mKey As String

' STEP 1: detect valid months
For r = 2 To lastRow
    If wsNonLocal.Cells(r, "C").Value = millerName _
    And wsNonLocal.Cells(r, "B").Value <= cutoffDate Then

        mKey = Format(wsNonLocal.Cells(r, "B").Value, "YYYY-MM")

        If Not dictMonth.Exists(mKey) Then
            dictMonth.Add mKey, False
        End If

        If Val(wsNonLocal.Cells(r, "T").Value) > 0 Then
            dictMonth(mKey) = True
        End If

    End If
Next r

' STEP 2: output only valid months
With wsNonLocal
    For r = 2 To lastRow
        If .Cells(r, "C").Value = millerName _
        And .Cells(r, "B").Value <= cutoffDate Then

            mKey = Format(.Cells(r, "B").Value, "YYYY-MM")

            If dictMonth.Exists(mKey) Then
                If dictMonth(mKey) = True Then

                    dataFound = True
                    wsRes.Cells(outRow, 1).Value = .Cells(r, "B").Value
                    wsRes.Cells(outRow, 2).Value = .Cells(r, "D").Value
                    wsRes.Cells(outRow, 3).Value = .Cells(r, "G").Value
                    wsRes.Cells(outRow, 4).Value = .Cells(r, "Q").Value
                    wsRes.Cells(outRow, 5).Value = .Cells(r, "R").Value
                    wsRes.Cells(outRow, 6).Value = .Cells(r, "S").Value
                    wsRes.Cells(outRow, 7).Value = .Cells(r, "T").Value

                    outRow = outRow + 1

                End If
            End If

        End If
    Next r
End With

'========================

If dataFound = False Then
    MsgBox "Data for '" & millerName & "' not found"
    GoTo SafeExit
End If

wsRes.Range("A3:G" & outRow - 1).Sort Key1:=wsRes.Range("A4"), Order1:=xlAscending, Header:=xlYes
Set tbl = wsRes.ListObjects.Add(xlSrcRange, wsRes.Range("A3:G" & outRow - 1), , xlYes)
tbl.TableStyle = "TableStyleMedium2"
tbl.ShowTotals = True
tbl.ListColumns("DELIVERY QTY").TotalsCalculation = xlTotalsCalculationSum
tbl.ListColumns("M COM").TotalsCalculation = xlTotalsCalculationSum
tbl.ListColumns("M CREDIT").TotalsCalculation = xlTotalsCalculationSum
tbl.ListColumns("M DEBIT").TotalsCalculation = xlTotalsCalculationSum

wsRes.Range("A1:G" & outRow).Font.Bold = True
MsgBox "Data for '" & millerName & "' copied to Results_NONLOCAL"

SafeExit:
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


'========================
' FROM–TO DATE FIXED (SAFE VERSION)
'========================
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsNonLocal As Worksheet
    Dim r As Long, lastRow As Long
    Dim millerName As String
    Dim fromDate As Date, toDate As Date
    Dim dataFound As Boolean
    Dim baseVal As Double

    Set wsNonLocal = Sheets("NONLOCAL")

    If Not Intersect(Target, Me.Range("E:F")) Is Nothing Then
        
        If Target.Row = 1 Then GoTo SafeExit
        
        millerName = Me.Cells(Target.Row, "B").Value
        
        If Me.Cells(Target.Row, "E").Value = "" Or Me.Cells(Target.Row, "F").Value = "" Then GoTo SafeExit
        
        fromDate = Me.Cells(Target.Row, "E").Value
        toDate = Me.Cells(Target.Row, "F").Value
        
        lastRow = wsNonLocal.Cells(wsNonLocal.Rows.Count, "B").End(xlUp).Row
        dataFound = False
        
        With wsNonLocal
            For r = 2 To lastRow
                
                If .Cells(r, "C").Value = millerName Then
                    
                    If IsDate(.Cells(r, "B").Value) Then
                        
                        If Int(.Cells(r, "B").Value) >= fromDate And Int(.Cells(r, "B").Value) <= toDate Then
                            
                            dataFound = True
                            
                            baseVal = Val(.Cells(r, "T").Value)
                            .Cells(r, "S").Value = baseVal
                            .Cells(r, "T").Value = baseVal - Val(.Cells(r, "R").Value)
                            
                        End If
                        
                    End If
                    
                End If
                
            Next r
        End With
        
        If dataFound = False Then
            MsgBox "Data not found for '" & millerName & "'"
        Else
            MsgBox "Data updated for '" & millerName & "'"
        End If
        
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
