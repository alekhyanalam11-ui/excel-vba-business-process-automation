'======================================================================
' TRIGGER ON DATE ENTRY (COLS F & G) - REMAINS UNCHANGED
'======================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.CountLarge > 1 Then Exit Sub
    If Intersect(Target, Me.Range("F:G")) Is Nothing Then Exit Sub
    
    Dim wsLocal As Worksheet, wsNonLocal As Worksheet
    Dim partyName As String, fromDate As Variant, toDate As Variant
    Dim r As Long, lastRow As Long
    Dim matchCount As Long
    
    partyName = Me.Cells(Target.Row, "B").Value
    fromDate = Me.Cells(Target.Row, "F").Value
    toDate = Me.Cells(Target.Row, "G").Value
    
    If partyName = "" Or Not IsDate(fromDate) Or Not IsDate(toDate) Then Exit Sub
    
    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set wsLocal = ThisWorkbook.Sheets("LOCAL")
    Set wsNonLocal = ThisWorkbook.Sheets("NONLOCAL")
    matchCount = 0
    
    lastRow = wsLocal.Cells(wsLocal.Rows.Count, "B").End(xlUp).Row
    For r = 2 To lastRow
        If wsLocal.Cells(r, "D").Value = partyName And _
           wsLocal.Cells(r, "B").Value >= fromDate And _
           wsLocal.Cells(r, "B").Value <= toDate Then
            wsLocal.Cells(r, "P").Value = wsLocal.Cells(r, "Q").Value
            wsLocal.Cells(r, "Q").Value = Val(wsLocal.Cells(r, "O").Value) - Val(wsLocal.Cells(r, "P").Value)
            matchCount = matchCount + 1
        End If
    Next r
    
    '====================================================
' BUILD MONTHLY BALANCE FROM COLUMN N
'====================================================

Set monthBalance = CreateObject("Scripting.Dictionary")

Set monthOrder = New Collection
lastSettledMonth = ""
runningBalance = 0

lastRow = wsNonLocal.Cells(wsNonLocal.Rows.Count, "B").End(xlUp).Row

For r = 2 To lastRow

    If wsNonLocal.Cells(r, "D").Value = partyName _
    And wsNonLocal.Cells(r, "B").Value <= cutoffDate Then

        monthKey = Format(wsNonLocal.Cells(r, "B").Value, "yyyymm")

        If Not monthBalance.Exists(monthKey) Then

            monthBalance.Add monthKey, 0
            monthOrder.Add monthKey

        End If

        'Column N already contains P DEBIT
        monthBalance(monthKey) = monthBalance(monthKey) + _
                                 Val(wsNonLocal.Cells(r, "N").Value)

    End If

Next r

'====================================================
' FIND LAST FULLY SETTLED MONTH
'====================================================

runningBalance = 0
lastSettledMonth = ""

For Each k In monthOrder

    runningBalance = runningBalance + monthBalance(k)

    If Round(runningBalance, 2) = 0 Then

        lastSettledMonth = k

    End If

Next k
    
    If matchCount > 0 Then
        MsgBox "UPDATED: " & matchCount & " records modified.", vbInformation
    Else
        MsgBox "NO DATA FOUND", vbExclamation
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'======================================================================
' TRIGGER ON SELECTION (REPORTS & TOTALS) - MONTHLY LOGIC ADDED
'======================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.CountLarge > 1 Then Exit Sub
    If Target.Column <> 2 And Target.Address <> "$I$1" Then Exit Sub
    If Target.Row = 1 And Target.Address <> "$I$1" Then Exit Sub
    
    Dim wsLocal As Worksheet, wsRes As Worksheet
Dim wsNonLocal As Worksheet, wsResNon As Worksheet

Dim r As Long
Dim i As Long
Dim lastRow As Long
Dim outRow As Long
Dim lastPartyRow As Long

Dim cutoffDate As Date
Dim partyName As String
Dim sno As String
Dim key As String
Dim monthKey As String

Dim tbl As ListObject

Dim localFound As Boolean
Dim nonLocalFound As Boolean

Dim dict As Object

Dim monthDictLocal As Object
Dim monthDictNon As Object

Dim monthDebit As Object
Dim monthCredit As Object
Dim monthBalance As Object

Dim arrLocal
Dim arrNonLocal
Dim arrParty
Dim resultArr()

Dim amt As Double
Dim localTotal As Double
Dim nonLocalTotal As Double

Dim carryForward As Double
Dim runningBalance As Double

Dim lastSettledMonth As String
Dim monthOrder As Collection

Dim k As Variant

    On Error GoTo SafeExit
    Application.EnableEvents = False

    Set wsLocal = ThisWorkbook.Sheets("LOCAL")
    Set wsRes = ThisWorkbook.Sheets("Results_LOCAL")
    Set wsNonLocal = ThisWorkbook.Sheets("NONLOCAL")
    Set wsResNon = ThisWorkbook.Sheets("Results_NONLOCAL")
    cutoffDate = DateSerial(Year(Date), Month(Date), 0)

    ' SECTION 1: TOTAL DUE UPDATE (I1 Trigger)
    If Target.Address = "$I$1" Then
        Set dict = CreateObject("Scripting.Dictionary")
        lastRow = wsLocal.Cells(wsLocal.Rows.Count, "B").End(xlUp).Row
        If lastRow >= 2 Then
            arrLocal = wsLocal.Range("B2:Q" & lastRow).Value
            For i = 1 To UBound(arrLocal, 1)
                If IsDate(arrLocal(i, 1)) And arrLocal(i, 1) <= cutoffDate Then
                    key = Trim(arrLocal(i, 3)): amt = Val(arrLocal(i, 16))
                    If key <> "" Then dict(key) = dict(key) + amt
                End If
            Next i
        End If
        lastRow = wsNonLocal.Cells(wsNonLocal.Rows.Count, "B").End(xlUp).Row
        If lastRow >= 2 Then
            arrNonLocal = wsNonLocal.Range("B2:N" & lastRow).Value
            For i = 1 To UBound(arrNonLocal, 1)
                If IsDate(arrNonLocal(i, 1)) And arrNonLocal(i, 1) <= cutoffDate Then
                    key = Trim(arrNonLocal(i, 3)): amt = Val(arrNonLocal(i, 13))
                    If key <> "" Then dict(key) = dict(key) + amt
                End If
            Next i
        End If
        lastPartyRow = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
        If lastPartyRow >= 2 Then
            arrParty = Me.Range("B2:B" & lastPartyRow).Value
            ReDim resultArr(1 To UBound(arrParty, 1), 1 To 1)
            For i = 1 To UBound(arrParty, 1)
                key = Trim(arrParty(i, 1))
                resultArr(i, 1) = IIf(dict.Exists(key), dict(key), 0)
            Next i
            Me.Range("D2:D" & lastPartyRow).Value = resultArr
        End If
        MsgBox "Totals Updated", vbInformation
        GoTo SafeExit
    End If

    ' SECTION 2: INDIVIDUAL REPORT (Party Name Click)
    partyName = Target.Value
    sno = Me.Cells(Target.Row, "A").Value
    Set monthDictLocal = CreateObject("Scripting.Dictionary")
    Set monthDictNon = CreateObject("Scripting.Dictionary")

    '--- LOCAL REPORT LOGIC ---
    wsRes.Cells.Clear
    With wsRes.Range("Q1"): .Value = "CLICK HERE FOR SINGLE ENTRY": .Font.Bold = True: End With
    For Each tbl In wsRes.ListObjects: tbl.Delete: Next tbl

'====================================================
' BUILD MONTHLY BALANCE FROM COLUMN Q
'====================================================

Set monthBalance = CreateObject("Scripting.Dictionary")

Set monthOrder = New Collection
lastSettledMonth = ""
runningBalance = 0

lastRow = wsLocal.Cells(wsLocal.Rows.Count, "B").End(xlUp).Row

For r = 2 To lastRow

    If wsLocal.Cells(r, "D").Value = partyName _
    And wsLocal.Cells(r, "B").Value <= cutoffDate Then

        monthKey = Format(wsLocal.Cells(r, "B").Value, "yyyymm")

        If Not monthBalance.Exists(monthKey) Then

            monthBalance.Add monthKey, 0

            monthOrder.Add monthKey

        End If

        'Column Q already contains O - P
        monthBalance(monthKey) = monthBalance(monthKey) + _
                                 Val(wsLocal.Cells(r, "Q").Value)

    End If

Next r

'====================================================
' FIND LAST FULLY SETTLED MONTH
'====================================================

runningBalance = 0
lastSettledMonth = ""

For Each k In monthOrder

    runningBalance = runningBalance + monthBalance(k)

    If Round(runningBalance, 2) = 0 Then

        lastSettledMonth = k

    End If

Next k

'====================================================
' FIND LAST FULLY SETTLED MONTH
'====================================================

runningBalance = 0
lastSettledMonth = ""

For Each k In monthOrder

    runningBalance = runningBalance + monthBalance(k)

    If Round(runningBalance, 2) = 0 Then

        lastSettledMonth = k

    End If

Next k

    ' Step B: Headers
    With wsRes
        .Range("A1").Value = sno & " - " & partyName
        .Range("A1:H1").Merge: .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Font.Size = 18: .Range("A1").Font.Bold = True
        .Range("A3").Resize(1, 10).Value = Array("DL DATE", "MILLER", "DELIVERY QTY", "CR MILLER", "CR PARTY", "M COM", "P COM", "TOTAL", "CREDIT", "DEBIT")
        .Rows(3).Font.Bold = True
    End With

    ' Step C: Transfer data only if Month total > 0
    outRow = 4

For r = 2 To lastRow

    If wsLocal.Cells(r, "D").Value = partyName _
    And wsLocal.Cells(r, "B").Value <= cutoffDate Then

        monthKey = Format(wsLocal.Cells(r, "B").Value, "yyyymm")

       If monthBalance.Exists(monthKey) Then

    'Skip every month up to the last fully settled month
    If lastSettledMonth = "" Or monthKey > lastSettledMonth Then

        If Round(monthBalance(monthKey), 2) <> 0 Then

            wsRes.Cells(outRow, 1).Resize(1, 10).Value = Array( _
                wsLocal.Cells(r, "B").Value, _
                wsLocal.Cells(r, "C").Value, _
                wsLocal.Cells(r, "G").Value, _
                wsLocal.Cells(r, "M").Value, _
                wsLocal.Cells(r, "K").Value, _
                wsLocal.Cells(r, "N").Value, _
                wsLocal.Cells(r, "L").Value, _
                wsLocal.Cells(r, "O").Value, _
                wsLocal.Cells(r, "P").Value, _
                wsLocal.Cells(r, "Q").Value)

            outRow = outRow + 1
            localFound = True

        End If

    End If

End If

    End If

Next r

'====================================================
' IF EVERYTHING IS SETTLED, DON'T DISPLAY OLD DATA
'====================================================

If lastSettledMonth <> "" Then

    If outRow = 4 Then

        localFound = False

    End If

End If

    If localFound Then
        wsRes.Range("A3:J" & outRow - 1).Sort Key1:=wsRes.Range("A4"), Order1:=xlAscending, Header:=xlYes
        Set tbl = wsRes.ListObjects.Add(xlSrcRange, wsRes.Range("A3:J" & outRow - 1), , xlYes)
        tbl.TableStyle = "TableStyleMedium2": tbl.ShowTotals = True: wsRes.Range("A3:J" & outRow).Font.Bold = True
        tbl.ListColumns("DELIVERY QTY").TotalsCalculation = xlTotalsCalculationSum
        tbl.ListColumns("TOTAL").TotalsCalculation = xlTotalsCalculationSum
        tbl.ListColumns("CREDIT").TotalsCalculation = xlTotalsCalculationSum
        tbl.ListColumns("DEBIT").TotalsCalculation = xlTotalsCalculationSum
        If Not tbl.TotalsRowRange Is Nothing Then tbl.TotalsRowRange.Font.Bold = True
    End If

    '--- NONLOCAL REPORT LOGIC ---
    wsResNon.Cells.Clear
    With wsResNon.Range("Q1"): .Value = "CLICK HERE FOR SINGLE ENTRY": .Font.Bold = True: End With
    For Each tbl In wsResNon.ListObjects: tbl.Delete: Next tbl
    
'====================================================
' BUILD MONTHLY BALANCE FROM NONLOCAL COLUMN N
'====================================================

Set monthBalance = CreateObject("Scripting.Dictionary")

Set monthOrder = New Collection
lastSettledMonth = ""
runningBalance = 0

lastRow = wsNonLocal.Cells(wsNonLocal.Rows.Count, "B").End(xlUp).Row

For r = 2 To lastRow

    If Trim(wsNonLocal.Cells(r, "D").Value) = Trim(partyName) _
    And IsDate(wsNonLocal.Cells(r, "B").Value) _
    And wsNonLocal.Cells(r, "B").Value <= cutoffDate Then

        monthKey = Format(wsNonLocal.Cells(r, "B").Value, "yyyymm")

        If Not monthBalance.Exists(monthKey) Then
            monthBalance.Add monthKey, 0
            monthOrder.Add monthKey
        End If

        monthBalance(monthKey) = monthBalance(monthKey) + _
                                 Val(wsNonLocal.Cells(r, "N").Value)

    End If

Next r

'====================================================
' FIND LAST FULLY SETTLED MONTH
'====================================================

runningBalance = 0
lastSettledMonth = ""

For Each k In monthOrder

    runningBalance = runningBalance + monthBalance(k)

    If Round(runningBalance, 2) = 0 Then
        lastSettledMonth = k
    End If

Next k
    With wsResNon
        .Range("A1").Value = sno & " - " & partyName
        .Range("A1:E1").Merge: .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").Font.Size = 18: .Range("A1").Font.Bold = True
        .Range("A3").Resize(1, 7).Value = Array("DL DATE", "MILLER", "DELIVERY QTY", "CR PARTY", "P COM", "P CREDIT", "P DEBIT")
        .Rows(3).Font.Bold = True
    End With

    outRow = 4

For r = 2 To lastRow

    If wsNonLocal.Cells(r, "D").Value = partyName _
    And wsNonLocal.Cells(r, "B").Value <= cutoffDate Then

        monthKey = Format(wsNonLocal.Cells(r, "B").Value, "yyyymm")

        If monthBalance.Exists(monthKey) Then

            'Skip every month up to the last fully settled month
            If lastSettledMonth = "" Or monthKey > lastSettledMonth Then

                'Keep month if its balance is not exactly zero
                If Round(monthBalance(monthKey), 2) <> 0 Then

                    wsResNon.Cells(outRow, 1).Resize(1, 7).Value = Array( _
                        wsNonLocal.Cells(r, "B").Value, _
                        wsNonLocal.Cells(r, "C").Value, _
                        wsNonLocal.Cells(r, "G").Value, _
                        wsNonLocal.Cells(r, "K").Value, _
                        wsNonLocal.Cells(r, "L").Value, _
                        wsNonLocal.Cells(r, "M").Value, _
                        wsNonLocal.Cells(r, "N").Value)

                    outRow = outRow + 1
                    nonLocalFound = True

                End If

            End If

        End If

    End If

Next r

'====================================================
' IF EVERYTHING IS SETTLED, DON'T DISPLAY OLD DATA
'====================================================

If lastSettledMonth <> "" Then

    If outRow = 4 Then

        nonLocalFound = False

    End If

End If

    If nonLocalFound Then
        wsResNon.Range("A3:G" & outRow - 1).Sort Key1:=wsResNon.Range("A4"), Order1:=xlAscending, Header:=xlYes
        Set tbl = wsResNon.ListObjects.Add(xlSrcRange, wsResNon.Range("A3:G" & outRow - 1), , xlYes)
        tbl.TableStyle = "TableStyleMedium2": tbl.ShowTotals = True: wsResNon.Range("A3:G" & outRow).Font.Bold = True
        tbl.ListColumns("DELIVERY QTY").TotalsCalculation = xlTotalsCalculationSum
        tbl.ListColumns(5).TotalsCalculation = xlTotalsCalculationSum
        tbl.ListColumns(6).TotalsCalculation = xlTotalsCalculationSum
        tbl.ListColumns(7).TotalsCalculation = xlTotalsCalculationSum
        If Not tbl.TotalsRowRange Is Nothing Then tbl.TotalsRowRange.Font.Bold = True
    End If

    If localFound Or nonLocalFound Then MsgBox "DATA FOUND FOR PARTY" Else MsgBox "NO DATA"

SafeExit:
    Application.EnableEvents = True
End Sub
