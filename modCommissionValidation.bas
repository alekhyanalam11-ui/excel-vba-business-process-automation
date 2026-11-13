Sub CheckAllCommissionRates()

    Dim wsParty As Worksheet
    Dim ws As Worksheet
    Dim dictParty As Object

    Dim lastRowParty As Long
    Dim lastRow As Long
    Dim i As Long

    Dim partyName As String
    Dim expectedRate As Variant

    Set wsParty = ThisWorkbook.Worksheets("PARTY")
    Set dictParty = CreateObject("Scripting.Dictionary")

    '=========================
    ' Build PARTY Dictionary
    '=========================
    lastRowParty = wsParty.Cells(wsParty.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRowParty

        partyName = Trim(CStr(wsParty.Cells(i, "B").Value))

        If partyName <> "" Then
            dictParty(partyName) = wsParty.Cells(i, "C").Value
        End If

    Next i

    '=========================
    ' CHECK LOCAL SHEET
    '=========================
    Set ws = ThisWorkbook.Worksheets("LOCAL")

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow

        partyName = Trim(CStr(ws.Cells(i, "D").Value))

        If partyName <> "" Then

            If dictParty.Exists(partyName) Then

                expectedRate = dictParty(partyName)

                If CDbl(Val(ws.Cells(i, "K").Value)) <> CDbl(Val(expectedRate)) Then
                    ws.Cells(i, "K").Interior.Color = RGB(255, 255, 0)
                Else
                    ws.Cells(i, "K").Interior.Pattern = xlNone
                End If

            Else
                ws.Cells(i, "K").Interior.Pattern = xlNone
            End If

        Else
            ws.Cells(i, "K").Interior.Pattern = xlNone
        End If

    Next i

    '=========================
    ' CHECK NONLOCAL SHEET
    '=========================
    Set ws = ThisWorkbook.Worksheets("NONLOCAL")

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow

        partyName = Trim(CStr(ws.Cells(i, "D").Value))

        If partyName <> "" Then

            If dictParty.Exists(partyName) Then

                expectedRate = dictParty(partyName)

                If CDbl(Val(ws.Cells(i, "K").Value)) <> CDbl(Val(expectedRate)) Then
                    ws.Cells(i, "K").Interior.Color = RGB(255, 255, 0)
                Else
                    ws.Cells(i, "K").Interior.Pattern = xlNone
                End If

            Else
                ws.Cells(i, "K").Interior.Pattern = xlNone
            End If

        Else
            ws.Cells(i, "K").Interior.Pattern = xlNone
        End If

    Next i

    MsgBox "Commission rate verification completed for LOCAL and NONLOCAL.", vbInformation

End Sub
