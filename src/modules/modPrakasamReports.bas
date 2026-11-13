Public Sub Extract_Prakasham_Pending_Delivery()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim i As Long
    Dim partyName As String
    Dim partyFirstWord As String
    Dim includeRow As Boolean
    Dim tbl As ListObject
    Dim outputRange As Range
    Dim lastDestRow As Long

    Set wsSource = ThisWorkbook.Worksheets("NONLOCAL")
    Set wsDest = ThisWorkbook.Worksheets("PRAKASAM")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    wsDest.ListObjects("tblPrakasamPendingDelivery").Delete
    wsDest.ListObjects("tblPrakasamPendingCommission").Delete
    On Error GoTo 0

    'Clear previous output
    wsDest.Range("A3:F" & wsDest.Rows.Count).Clear

    'Headers
    wsDest.Range("A3").Value = wsSource.Range("A1").Value
    wsDest.Range("B3").Value = wsSource.Range("B1").Value
    wsDest.Range("C3").Value = wsSource.Range("C1").Value
    wsDest.Range("D3").Value = wsSource.Range("D1").Value
    wsDest.Range("E3").Value = wsSource.Range("F1").Value
    wsDest.Range("F3").Value = wsSource.Range("G1").Value

    destRow = 4

    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow

        partyName = Trim(CStr(wsSource.Cells(i, "D").Value))

        If partyName <> "" Then

            partyFirstWord = UCase(Split(partyName, " ")(0))
            includeRow = False

            Select Case partyFirstWord

                Case "PAMUR", "PAMURU", "PODILI", "KANIGIRI", _
                     "KANDUKUR", "SINGHARAYAKONDA", "CHIMAKURTHI", _
                     "SANTANOOTALAPADU", "ADDANKI", "ONG", "MADDIPADU", "TANGUTURU"

                    includeRow = True

            End Select

            'PARTY condition + DL DATE blank
            If includeRow = True Then

                If Trim(CStr(wsSource.Cells(i, "B").Value)) = "" Then

                    wsDest.Cells(destRow, "A").Value = wsSource.Cells(i, "A").Value
                    wsDest.Cells(destRow, "B").Value = wsSource.Cells(i, "B").Value
                    wsDest.Cells(destRow, "C").Value = wsSource.Cells(i, "C").Value
                    wsDest.Cells(destRow, "D").Value = wsSource.Cells(i, "D").Value
                    wsDest.Cells(destRow, "E").Value = wsSource.Cells(i, "F").Value
                    wsDest.Cells(destRow, "F").Value = wsSource.Cells(i, "G").Value

                    destRow = destRow + 1

                End If

            End If

        End If

    Next i

    lastDestRow = destRow - 1

    If lastDestRow >= 4 Then

        Set outputRange = wsDest.Range("A3:F" & lastDestRow)

        Set tbl = wsDest.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=outputRange, _
                    XlListObjectHasHeaders:=xlYes)

        tbl.Name = "tblPrakasamPendingDelivery"
        tbl.TableStyle = "TableStyleMedium2"

        'Sort by MILLER column - A to Z
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsDest.Range("C4:C" & lastDestRow), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With

        'Bold headers and data
        wsDest.Range("A3:F" & lastDestRow).Font.Bold = True

    Else

        wsDest.Range("A3:F3").Font.Bold = True

    End If

    wsDest.Columns("A:F").AutoFit

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


Public Sub Extract_Prakasham_Pending_Commission()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim i As Long
    Dim partyName As String
    Dim partyFirstWord As String
    Dim includeRow As Boolean
    Dim pDebit As Variant
    Dim tbl As ListObject
    Dim outputRange As Range
    Dim lastDestRow As Long

    Set wsSource = ThisWorkbook.Worksheets("NONLOCAL")
    Set wsDest = ThisWorkbook.Worksheets("PRAKASAM")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    wsDest.ListObjects("tblPrakasamPendingDelivery").Delete
    wsDest.ListObjects("tblPrakasamPendingCommission").Delete
    On Error GoTo 0

    'Clear previous output
    wsDest.Range("A3:G" & wsDest.Rows.Count).Clear

    'Headers
    wsDest.Range("A3").Value = wsSource.Range("B1").Value
    wsDest.Range("B3").Value = wsSource.Range("C1").Value
    wsDest.Range("C3").Value = wsSource.Range("D1").Value
    wsDest.Range("D3").Value = wsSource.Range("G1").Value
    wsDest.Range("E3").Value = wsSource.Range("L1").Value
    wsDest.Range("F3").Value = wsSource.Range("M1").Value
    wsDest.Range("G3").Value = wsSource.Range("N1").Value

    destRow = 4

    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow

        partyName = Trim(CStr(wsSource.Cells(i, "D").Value))

        If partyName <> "" Then

            partyFirstWord = UCase(Split(partyName, " ")(0))
            includeRow = False

            Select Case partyFirstWord

                Case "PAMUR", "PAMURU", "PODILI", "KANIGIRI", _
                     "KANDUKUR", "SINGHARAYAKONDA", "CHIMAKURTHI", _
                     "SANTANOOTALAPADU", "ADDANKI", "ONG", "MADDIPADU", "TANGUTURU"

                    includeRow = True

            End Select

            If includeRow = True Then

                pDebit = wsSource.Cells(i, "N").Value

                'Keep only rows where DL DATE is a valid date
                'Text-format DL DATE rows are excluded
                If IsDate(wsSource.Cells(i, "B").Value) And _
                   Not IsEmpty(wsSource.Cells(i, "B").Value) Then

                    'Keep only P DEBIT values greater than 0
                    'Blank P DEBIT rows are excluded
                    If IsNumeric(pDebit) Then

                        If CDbl(pDebit) > 0 Then

                            wsDest.Cells(destRow, "A").Value = wsSource.Cells(i, "B").Value
                            wsDest.Cells(destRow, "B").Value = wsSource.Cells(i, "C").Value
                            wsDest.Cells(destRow, "C").Value = wsSource.Cells(i, "D").Value
                            wsDest.Cells(destRow, "D").Value = wsSource.Cells(i, "G").Value
                            wsDest.Cells(destRow, "E").Value = wsSource.Cells(i, "L").Value
                            wsDest.Cells(destRow, "F").Value = wsSource.Cells(i, "M").Value
                            wsDest.Cells(destRow, "G").Value = wsSource.Cells(i, "N").Value

                            destRow = destRow + 1

                        End If

                    End If

                End If

            End If

        End If

    Next i

    lastDestRow = destRow - 1

    If lastDestRow >= 4 Then

        Set outputRange = wsDest.Range("A3:G" & lastDestRow)

        Set tbl = wsDest.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=outputRange, _
                    XlListObjectHasHeaders:=xlYes)

        tbl.Name = "tblPrakasamPendingCommission"
        tbl.TableStyle = "TableStyleMedium2"

        'Sort by MILLER column - A to Z
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsDest.Range("B4:B" & lastDestRow), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With

        'Bold headers and data
        wsDest.Range("A3:G" & lastDestRow).Font.Bold = True

    Else

        wsDest.Range("A3:G3").Font.Bold = True

    End If

    wsDest.Columns("A:G").AutoFit

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
