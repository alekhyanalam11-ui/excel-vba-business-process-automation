Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Address <> "$H$1" Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False

    Dim wsPat As Worksheet
    Dim wsLocal As Worksheet
    Dim wsNonLocal As Worksheet

    Dim lastRowPat As Long
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

    Set wsPat = Me
    Set wsLocal = Sheets("LOCAL")
    Set wsNonLocal = Sheets("NONLOCAL")

    cutoffDate = DateSerial(Year(Date), Month(Date), 0)

    lastRowPat = wsPat.Cells(wsPat.Rows.Count, "B").End(xlUp).Row
    lastRowLocal = wsLocal.Cells(wsLocal.Rows.Count, "D").End(xlUp).Row
    lastRowNonLocal = wsNonLocal.Cells(wsNonLocal.Rows.Count, "D").End(xlUp).Row

    '====================================
    'REMOVE OLD TOTAL ROWS (NO DUPLICATE)
    '====================================
    For i = lastRowPat To 4 Step -1
        If wsPat.Cells(i, "B").Value = "TOTAL SUM" Or _
           wsPat.Cells(i, "B").Value = "TOTAL COUNT" Then
            wsPat.Range("B" & i & ":C" & i).ClearContents
        End If
    Next i

    '====================================
    'CLEAR OLD VALUES IN COLUMN C
    '====================================
    wsPat.Range("C4:C" & lastRowPat).ClearContents

    '====================================
    'MAIN CALCULATION
    '====================================
    For i = 4 To lastRowPat

        partyName = Trim(wsPat.Cells(i, "B").Value)
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
            With wsPat.Cells(i, "C")
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
            wsPat.Cells(i, "C").ClearContents
        End If

    Next i

    '====================================
    'FIND TRUE LAST DATA ROW
    '====================================
    trueLastRow = 4

    For i = lastRowPat To 4 Step -1
        If Trim(wsPat.Cells(i, "B").Value) <> "" Then
            trueLastRow = i
            Exit For
        End If
    Next i

    '====================================
    'PLACE SUMMARY (WITH GAP)
    '====================================

    outputRow = trueLastRow + 2   '1 blank row

    'TOTAL SUM
    wsPat.Cells(outputRow, "B").Value = "TOTAL SUM"

    If sumTotal <> 0 Then
        wsPat.Cells(outputRow, "C").Value = sumTotal
    Else
        wsPat.Cells(outputRow, "C").ClearContents
    End If

    wsPat.Cells(outputRow, "B").Font.Bold = True
    wsPat.Cells(outputRow, "C").Font.Bold = True

    '1 blank row
    outputRow = outputRow + 2

    'TOTAL COUNT
    wsPat.Cells(outputRow, "B").Value = "TOTAL COUNT"

    If countTotal <> 0 Then
        wsPat.Cells(outputRow, "C").Value = countTotal
    Else
        wsPat.Cells(outputRow, "C").ClearContents
    End If

    wsPat.Cells(outputRow, "B").Font.Bold = True
    wsPat.Cells(outputRow, "C").Font.Bold = True

SafeExit:
    Application.EnableEvents = True

End Sub
