Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Address <> "$H$1" Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False

    Dim wsTnl As Worksheet
    Dim wsLocal As Worksheet
    Dim wsNonLocal As Worksheet

    Dim lastRowTnl As Long
    Dim lastRowLocal As Long
    Dim lastRowNonLocal As Long

    Dim i As Long, j As Long
    Dim partyName As String
    Dim totalDue As Double
    Dim cutoffDate As Date

    Dim sumTotal As Double
    Dim countTotal As Long
    Dim outputRow As Long
    Dim trueLastRow As Long

    Set wsTnl = Me
    Set wsLocal = Sheets("LOCAL")
    Set wsNonLocal = Sheets("NONLOCAL")

    cutoffDate = DateSerial(Year(Date), Month(Date), 0)

    lastRowTnl = wsTnl.Cells(wsTnl.Rows.Count, "B").End(xlUp).Row
    lastRowLocal = wsLocal.Cells(wsLocal.Rows.Count, "D").End(xlUp).Row
    lastRowNonLocal = wsNonLocal.Cells(wsNonLocal.Rows.Count, "D").End(xlUp).Row

    '====================================
    'REMOVE OLD TOTAL ROWS (NO DUPLICATE)
    '====================================
    For i = lastRowTnl To 4 Step -1
        If wsTnl.Cells(i, "B").Value = "TOTAL SUM" Or _
           wsTnl.Cells(i, "B").Value = "TOTAL COUNT" Then
            wsTnl.Range("B" & i & ":C" & i).ClearContents
        End If
    Next i

    '====================================
    'CLEAR OLD VALUES IN COLUMN C
    '====================================
    wsTnl.Range("C4:C" & lastRowTnl).ClearContents

    '====================================
    'MAIN CALCULATION
    '====================================
    For i = 4 To lastRowTnl

        partyName = Trim(wsTnl.Cells(i, "B").Value)
        totalDue = 0

        If partyName <> "" Then

            'LOCAL
            For j = 2 To lastRowLocal
                If Trim(wsLocal.Cells(j, "D").Value) = partyName Then
                    If IsDate(wsLocal.Cells(j, "B").Value) Then
                        If wsLocal.Cells(j, "B").Value <= cutoffDate Then
                            totalDue = totalDue + Val(wsLocal.Cells(j, "Q").Value)
                        End If
                    End If
                End If
            Next j

            'NONLOCAL
            For j = 2 To lastRowNonLocal
                If Trim(wsNonLocal.Cells(j, "D").Value) = partyName Then
                    If IsDate(wsNonLocal.Cells(j, "B").Value) Then
                        If wsNonLocal.Cells(j, "B").Value <= cutoffDate Then
                            totalDue = totalDue + Val(wsNonLocal.Cells(j, "N").Value)
                        End If
                    End If
                End If
            Next j

            '?? WRITE ONLY IF VALUE EXISTS
            With wsTnl.Cells(i, "C")
                If totalDue <> 0 Then
                    .Value = totalDue
                Else
                    .ClearContents
                End If
                .Font.Bold = True
            End With

            sumTotal = sumTotal + totalDue

            If totalDue <> 0 Then
                countTotal = countTotal + 1
            End If

        Else
            '?? EMPTY ROW ? KEEP BLANK (NO ZERO)
            wsTnl.Cells(i, "C").ClearContents
        End If

    Next i

    '====================================
    'FIND TRUE LAST DATA ROW
    '====================================
    trueLastRow = 4

    For i = lastRowTnl To 4 Step -1
        If Trim(wsTnl.Cells(i, "B").Value) <> "" Then
            trueLastRow = i
            Exit For
        End If
    Next i

    '====================================
    'PLACE SUMMARY (WITH GAP)
    '====================================

    outputRow = trueLastRow + 2   '1 blank row

    'TOTAL SUM
    wsTnl.Cells(outputRow, "B").Value = "TOTAL SUM"

    If sumTotal <> 0 Then
        wsTnl.Cells(outputRow, "C").Value = sumTotal
    Else
        wsTnl.Cells(outputRow, "C").ClearContents
    End If

    wsTnl.Cells(outputRow, "B").Font.Bold = True
    wsTnl.Cells(outputRow, "C").Font.Bold = True

    '1 blank row
    outputRow = outputRow + 2

    'TOTAL COUNT
    wsTnl.Cells(outputRow, "B").Value = "TOTAL COUNT"

    If countTotal <> 0 Then
        wsTnl.Cells(outputRow, "C").Value = countTotal
    Else
        wsTnl.Cells(outputRow, "C").ClearContents
    End If

    wsTnl.Cells(outputRow, "B").Font.Bold = True
    wsTnl.Cells(outputRow, "C").Font.Bold = True

SafeExit:
    Application.EnableEvents = True

End Sub
