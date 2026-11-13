'=========================================================
' K1 - LOCAL PENDING DELIVERY
'=========================================================
Public Sub UpdatePendingLocal()
    RefreshPendingDelivery "LOCAL", False
End Sub

'=========================================================
' K3 - NONLOCAL PENDING DELIVERY
'=========================================================
Public Sub UpdatePendingNonLocal()
    RefreshPendingDelivery "NONLOCAL", True
End Sub

'=========================================================
' CLEAR PREVIOUS AUTOMATION OUTPUT
' Preserves rows 1 and 2.
' Clears output and formatting in columns A:F.
'=========================================================
Public Sub ClearAutomationOutput()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastCell As Range
    Dim lastRow As Long
    Dim tableLastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("PENDING_DELIVERY")

    lastRow = 2

    ' Find the last used row in the output columns.
    Set lastCell = ws.Range("A:F").Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If Not lastCell Is Nothing Then
        lastRow = lastCell.Row
    End If

    ' Include old table rows even if their cells are empty.
    For i = ws.ListObjects.Count To 1 Step -1

        Set lo = ws.ListObjects(i)

        If Not Intersect(lo.Range, ws.Range("A:F")) Is Nothing Then

            If lo.Range.Row + lo.Range.Rows.Count - 1 >= 3 Then

                tableLastRow = lo.Range.Row + lo.Range.Rows.Count - 1

                If tableLastRow > lastRow Then
                    lastRow = tableLastRow
                End If

                ' Remove the old table object.
                lo.Unlist

            End If

        End If

    Next i

    ' Remove previous output and its formatting.
    If lastRow >= 3 Then
        ws.Range("A3:F" & lastRow).Clear
    End If

End Sub

'=========================================================
' COMMON REFRESH PROCEDURE FOR K1 AND K3
'=========================================================
Private Sub RefreshPendingDelivery( _
    ByVal sourceSheetName As String, _
    ByVal excludeParties As Boolean)

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastCell As Range
    Dim lastRow As Long
    Dim outputCount As Long
    Dim sourceData As Variant
    Dim outputData() As Variant
    Dim headers As Variant
    Dim excludeList As Variant
    Dim prefix As Variant
    Dim partyName As String
    Dim isExcluded As Boolean
    Dim i As Long
    Dim j As Long
    Dim lo As ListObject
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo ErrorHandler

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set wsSource = ThisWorkbook.Worksheets(sourceSheetName)
    Set wsDest = ThisWorkbook.Worksheets("PENDING_DELIVERY")

    ' Remove the previous automation output.
    ClearAutomationOutput

    ' Headers in row 3.
    headers = Array("OR DATE", "DL DATE", "MILLER", _
                    "PARTY", "ORDER QTY", "DELIVERY QTY")

    For j = 0 To 5
        wsDest.Cells(3, j + 1).Value = headers(j)
    Next j

    ' Find the last used row in source columns A:G.
    Set lastCell = wsSource.Range("A:G").Find( _
        What:="*", _
        After:=wsSource.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastCell Is Nothing Then GoTo FormatOutput

    lastRow = lastCell.Row

    If lastRow < 2 Then GoTo FormatOutput

    sourceData = wsSource.Range("A1:G" & lastRow).Value2

    ' Excluded prefixes for NONLOCAL K3.
    excludeList = Array( _
        "PAMUR", "PAMURU", "PODILI", "KANIGIRI", _
        "KANDUKUR", "SINGHARAYAKONDA", "CHIMAKURTHI", _
        "SANTANOOTALAPADU", "ADDANKI", "ONG", "MADDIPADU", "TANGUTURU")

    ' Count matching rows.
    For i = 2 To UBound(sourceData, 1)

        If Not IsError(sourceData(i, 2)) Then

            If Len(Trim$(CStr(sourceData(i, 2)))) = 0 Then

                isExcluded = False

                If excludeParties Then

                    If Not IsError(sourceData(i, 4)) Then
                        partyName = UCase$(Trim$(CStr(sourceData(i, 4))))
                    Else
                        partyName = vbNullString
                    End If

                    For Each prefix In excludeList

                        If Left$(partyName, Len(CStr(prefix))) = CStr(prefix) Then
                            isExcluded = True
                            Exit For
                        End If

                    Next prefix

                End If

                If Not isExcluded Then
                    outputCount = outputCount + 1
                End If

            End If

        End If

    Next i

    If outputCount = 0 Then GoTo FormatOutput

    ReDim outputData(1 To outputCount, 1 To 6)

    outputCount = 0

    ' Copy source columns A, B, C, D, F, G.
    For i = 2 To UBound(sourceData, 1)

        If Not IsError(sourceData(i, 2)) Then

            If Len(Trim$(CStr(sourceData(i, 2)))) = 0 Then

                isExcluded = False

                If excludeParties Then

                    If Not IsError(sourceData(i, 4)) Then
                        partyName = UCase$(Trim$(CStr(sourceData(i, 4))))
                    Else
                        partyName = vbNullString
                    End If

                    For Each prefix In excludeList

                        If Left$(partyName, Len(CStr(prefix))) = CStr(prefix) Then
                            isExcluded = True
                            Exit For
                        End If

                    Next prefix

                End If

                If Not isExcluded Then

                    outputCount = outputCount + 1

                    outputData(outputCount, 1) = sourceData(i, 1)
                    outputData(outputCount, 2) = sourceData(i, 2)
                    outputData(outputCount, 3) = sourceData(i, 3)
                    outputData(outputCount, 4) = sourceData(i, 4)
                    outputData(outputCount, 5) = sourceData(i, 6)
                    outputData(outputCount, 6) = sourceData(i, 7)

                End If

            End If

        End If

    Next i

    ' Data starts from row 4.
    wsDest.Range("A4").Resize(outputCount, 6).Value = outputData

    ' Display dates instead of serial numbers.
    wsDest.Range("A4:A" & outputCount + 3).NumberFormat = "dd-mm-yyyy"

    ' Sort by column C in ascending order.
    With wsDest.Sort

        .SortFields.Clear

        .SortFields.Add _
            key:=wsDest.Range("C4:C" & outputCount + 3), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange wsDest.Range("A3:F" & outputCount + 3)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

FormatOutput:

    If outputCount > 0 Then

        Set lo = wsDest.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=wsDest.Range("A3:F" & outputCount + 3), _
            XlListObjectHasHeaders:=xlYes)

        lo.Name = "tblPendingDelivery"
        lo.TableStyle = "TableStyleMedium2"
        lo.Range.Font.Bold = True

    Else

        With wsDest.Range("A3:F3")
            .Font.Bold = True
            .Interior.Color = RGB(91, 155, 213)
        End With

    End If

    wsDest.Columns("A").NumberFormat = "dd-mm-yyyy"
    wsDest.Columns("A:F").AutoFit

CleanExit:

    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

ErrorHandler:

    MsgBox "Error updating " & sourceSheetName & _
           " pending deliveries:" & vbCrLf & Err.Description, _
           vbExclamation, "Pending Delivery"

    Resume CleanExit

End Sub

'=========================================================
' COMMON PROCEDURE FOR ALL SIX COLUMN L AUTOMATIONS
'=========================================================
Private Sub ExtractByPartyPrefixes(ByVal prefixGroup As String)

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim lastDestRow As Long
    Dim destRow As Long
    Dim i As Long
    Dim partyName As String
    Dim partyFirstWord As String
    Dim includeRow As Boolean
    Dim tbl As ListObject
    Dim outputRange As Range
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo ErrorHandler

    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set wsSource = ThisWorkbook.Worksheets("NONLOCAL")
    Set wsDest = ThisWorkbook.Worksheets("PENDING_DELIVERY")

    ' Clear the previous automation completely.
    ClearAutomationOutput

    ' Preserve the original headers from NONLOCAL.
    wsDest.Range("A3").Value = wsSource.Range("A1").Value
    wsDest.Range("B3").Value = wsSource.Range("B1").Value
    wsDest.Range("C3").Value = wsSource.Range("C1").Value
    wsDest.Range("D3").Value = wsSource.Range("D1").Value
    wsDest.Range("E3").Value = wsSource.Range("F1").Value
    wsDest.Range("F3").Value = wsSource.Range("G1").Value

    destRow = 4

    lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow

        If Not IsError(wsSource.Cells(i, "D").Value) Then

            partyName = Trim$(CStr(wsSource.Cells(i, "D").Value))

            If partyName <> "" Then

                partyFirstWord = UCase$(Split(partyName, " ")(0))
                includeRow = False

                Select Case prefixGroup

                    Case "KALI"
                        Select Case partyFirstWord
                            Case "ELR", "AGP", "JANGAREDDYGUDEM", _
                                 "NZV", "GOLLAPALLI", "GANNAVARAM", _
                                 "PEDAVEGI", "LAKKAVARAM", "VPT", "MYLAVARAM"
                                includeRow = True
                        End Select

                    Case "PRASAD"
                        Select Case partyFirstWord
                            Case "BHVM", "PKL", "GNT", _
                                 "KALIDINDI", "KAIKULUR", "AKV"
                                includeRow = True
                        End Select

                    Case "MADHU"
                        Select Case partyFirstWord
                            Case "MGR", "PEDAPALEM", "REVENDRAPADU", _
                                 "TADEPALLI", "NUTTAKKI", "STP", "MELLAMPUDI"
                                includeRow = True
                        End Select

                    Case "JEMILI"
                        Select Case partyFirstWord
                            Case "PNR", "KPM", "BPT"
                                includeRow = True
                        End Select

                    Case "MK"
                        Select Case partyFirstWord
                            Case "TNL", "BANDAR", "RPL", "CHPALLI", "MOPIDEVI"
                                includeRow = True
                        End Select

                    Case "ONG_VRR"
                        Select Case partyFirstWord
                            Case "KANDUKUR", "SINGHARAYAKONDA", "KAVALI"
                                includeRow = True
                        End Select

                End Select

                ' Preserve the original condition: DL DATE is blank.
                If includeRow Then

                    If Not IsError(wsSource.Cells(i, "B").Value) Then

                        If Trim$(CStr(wsSource.Cells(i, "B").Value)) = "" Then

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

        tbl.Name = "tblPendingDelivery"
        tbl.TableStyle = "TableStyleMedium2"

        ' Preserve original ascending sort by column C.
        With wsDest.Sort

            .SortFields.Clear

            .SortFields.Add _
                key:=wsDest.Range("C4:C" & lastDestRow), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal

            .SetRange wsDest.Range("A3:F" & lastDestRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply

        End With

        ' Bold all headers and data.
        wsDest.Range("A3:F" & lastDestRow).Font.Bold = True

    Else

        wsDest.Range("A3:F3").Font.Bold = True

    End If

    ' Display column A as dates.
    wsDest.Columns("A").NumberFormat = "dd-mm-yyyy"

    wsDest.Columns("A:F").AutoFit

CleanExit:

    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

ErrorHandler:

    MsgBox "Error in Extract_" & prefixGroup & ":" & _
           vbCrLf & Err.Description, _
           vbExclamation, "Pending Delivery"

    Resume CleanExit

End Sub

'=========================================================
' SIX ORIGINAL COLUMN L AUTOMATION PROCEDURES
'=========================================================
Public Sub Extract_KALI()
    ExtractByPartyPrefixes "KALI"
End Sub

Public Sub Extract_PRASAD()
    ExtractByPartyPrefixes "PRASAD"
End Sub

Public Sub Extract_MADHU()
    ExtractByPartyPrefixes "MADHU"
End Sub

Public Sub Extract_JEMILI()
    ExtractByPartyPrefixes "JEMILI"
End Sub

Public Sub Extract_MK()
    ExtractByPartyPrefixes "MK"
End Sub

Public Sub Extract_ONG_VRR()
    ExtractByPartyPrefixes "ONG_VRR"
End Sub
