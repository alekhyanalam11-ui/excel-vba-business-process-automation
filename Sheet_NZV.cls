Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Address <> "$H$1" Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim wsNzv As Worksheet
    Dim wsLocal As Worksheet
    Dim wsNonLocal As Worksheet

    Dim lastRowNzv As Long
    Dim lastRowLocal As Long
    Dim lastRowNonLocal As Long

    Dim i As Long
    Dim partyName As String
    Dim totalDue As Double
    Dim cutoffDate As Date

    Dim sumTotal As Double
    Dim countTotal As Long
    Dim outputRow As Long
    Dim trueLastRow As Long

    'DICTIONARIES
    Dim dictLocal As Object
    Dim dictNonLocal As Object
    Dim key As String

    Set dictLocal = CreateObject("Scripting.Dictionary")
    Set dictNonLocal = CreateObject("Scripting.Dictionary")

    Set wsNzv = Me
    Set wsLocal = Sheets("LOCAL")
    Set wsNonLocal = Sheets("NONLOCAL")

    cutoffDate = DateSerial(Year(Date), Month(Date), 0)

    lastRowNzv = wsNzv.Cells(wsNzv.Rows.Count, "B").End(xlUp).Row
    lastRowLocal = wsLocal.Cells(wsLocal.Rows.Count, "D").End(xlUp).Row
    lastRowNonLocal = wsNonLocal.Cells(wsNonLocal.Rows.Count, "D").End(xlUp).Row

    '====================================
    'BUILD LOCAL DICTIONARY
    '====================================
    For i = 2 To lastRowLocal

        If Trim(wsLocal.Cells(i, "D").Value) <> "" Then

            If IsDate(wsLocal.Cells(i, "B").Value) Then
                If wsLocal.Cells(i, "B").Value <= cutoffDate Then

                    key = Trim(wsLocal.Cells(i, "D").Value)
                    dictLocal(key) = dictLocal(key) + Val(wsLocal.Cells(i, "Q").Value)

                End If
            End If

        End If

    Next i

    '====================================
    'BUILD NONLOCAL DICTIONARY
    '====================================
    For i = 2 To lastRowNonLocal

        If Trim(wsNonLocal.Cells(i, "D").Value) <> "" Then

            If IsDate(wsNonLocal.Cells(i, "B").Value) Then
                If wsNonLocal.Cells(i, "B").Value <= cutoffDate Then

                    key = Trim(wsNonLocal.Cells(i, "D").Value)
                    dictNonLocal(key) = dictNonLocal(key) + Val(wsNonLocal.Cells(i, "N").Value)

                End If
            End If

        End If

    Next i

    '====================================
    'REMOVE OLD TOTAL ROWS
    '====================================
    For i = lastRowNzv To 4 Step -1
        If wsNzv.Cells(i, "B").Value = "TOTAL SUM" Or _
           wsNzv.Cells(i, "B").Value = "TOTAL COUNT" Then
            wsNzv.Range("B" & i & ":C" & i).ClearContents
        End If
    Next i

    '====================================
    'CLEAR OLD VALUES
    '====================================
    wsNzv.Range("C4:C" & lastRowNzv).ClearContents

    'RESET TOTALS
    sumTotal = 0
    countTotal = 0

    '====================================
    'MAIN CALCULATION (FAST LOOKUP)
    '====================================
    For i = 4 To lastRowNzv

        partyName = Trim(wsNzv.Cells(i, "B").Value)
        totalDue = 0

        If partyName <> "" Then

            If dictLocal.Exists(partyName) Then
                totalDue = totalDue + dictLocal(partyName)
            End If

            If dictNonLocal.Exists(partyName) Then
                totalDue = totalDue + dictNonLocal(partyName)
            End If

            With wsNzv.Cells(i, "C")
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
            wsNzv.Cells(i, "C").ClearContents
        End If

    Next i

    '====================================
    'FIND TRUE LAST DATA ROW
    '====================================
    trueLastRow = 4

    For i = lastRowNzv To 4 Step -1
        If Trim(wsNzv.Cells(i, "B").Value) <> "" Then
            trueLastRow = i
            Exit For
        End If
    Next i

    '====================================
    'SUMMARY
    '====================================
    outputRow = trueLastRow + 2

    wsNzv.Cells(outputRow, "B").Value = "TOTAL SUM"
    If sumTotal <> 0 Then
        wsNzv.Cells(outputRow, "C").Value = sumTotal
    End If

    wsNzv.Cells(outputRow, "B").Font.Bold = True
    wsNzv.Cells(outputRow, "C").Font.Bold = True

    outputRow = outputRow + 2

    wsNzv.Cells(outputRow, "B").Value = "TOTAL COUNT"
    If countTotal <> 0 Then
        wsNzv.Cells(outputRow, "C").Value = countTotal
    End If

    wsNzv.Cells(outputRow, "B").Font.Bold = True
    wsNzv.Cells(outputRow, "C").Font.Bold = True

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
