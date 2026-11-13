Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Cells.CountLarge > 1 Then Exit Sub

    Dim wsLoc As Worksheet, wsRes As Worksheet
    Dim lastRowMiller As Long, lastRowLocal As Long, i As Long, j As Long
    Dim millerName As String, cutoffDate As Date
    Dim totalDebitSum As Double

    On Error Resume Next
    Set wsLoc = Worksheets("LOCAL")
    On Error GoTo 0

    If wsLoc Is Nothing Then
        MsgBox "Sheet 'LOCAL' not found!", vbCritical
        Exit Sub
    End If

    cutoffDate = Date

    ' =========================
    ' ACTION 1 (UNCHANGED)
    ' =========================
    If Not Intersect(Target, Me.Range("K1")) Is Nothing Then

        Application.ScreenUpdating = False
        Application.EnableEvents = False

        lastRowMiller = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
        lastRowLocal = wsLoc.Cells(wsLoc.Rows.Count, "C").End(xlUp).Row

        If lastRowMiller >= 2 Then
            With Me.Range("D2:D" & lastRowMiller)
                .ClearContents
                .Font.Bold = True
            End With
        End If

        For i = 2 To lastRowMiller
            millerName = Trim(Me.Cells(i, "B").Value)
            totalDebitSum = 0

            If millerName <> "" Then
                For j = 2 To lastRowLocal
                    If Trim(wsLoc.Cells(j, "C").Value) = millerName Then
                        If IsDate(wsLoc.Cells(j, "B").Value) Then
                            If wsLoc.Cells(j, "B").Value < cutoffDate Then
                                totalDebitSum = totalDebitSum + Val(wsLoc.Cells(j, "Z").Value)
                            End If
                        End If
                    End If
                Next j

                Me.Cells(i, "D").Value = totalDebitSum
            End If
        Next i

        Application.EnableEvents = True
        Application.ScreenUpdating = True

        Me.Range("A1").Select

        MsgBox "Column D updated and BOLDED for dates before " & Format(cutoffDate, "DD/MM/YYYY"), vbInformation
        Exit Sub
    End If

    ' =========================
    ' ACTION 2 (UPDATED LOGIC)
    ' =========================
    If Target.Column <> 2 Or Target.Row = 1 Then Exit Sub

    On Error GoTo SafeExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set wsRes = Worksheets("Results_MILLER_LOCAL")
    millerName = Trim(Target.Value)
    If millerName = "" Then GoTo SafeExit

    wsRes.Cells.Clear

    With wsRes.Range("A1:E1")
        .Merge
        .Value = millerName
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    wsRes.Range("A3:G3").Value = Array("DL DATE", "PARTY", "DELIVERY QTY", "CR MILLER", "M COM", "M CREDIT", "M DEBIT")

    Dim outRow As Long, dataFound As Boolean
    outRow = 4
    dataFound = False
    lastRowLocal = wsLoc.Cells(wsLoc.Rows.Count, "C").End(xlUp).Row

    ' =========================
    ' ? MONTH FILTER (BASED ON Z)
    ' =========================
    Dim dictMonth As Object
    Set dictMonth = CreateObject("Scripting.Dictionary")

    Dim key As String

    ' STEP 1: detect valid months (any Z > 0)
    For i = 2 To lastRowLocal
        If Trim(wsLoc.Cells(i, "C").Value) = millerName Then
            If IsDate(wsLoc.Cells(i, "B").Value) Then
                If wsLoc.Cells(i, "B").Value < cutoffDate Then

                    key = Format(wsLoc.Cells(i, "B").Value, "YYYY-MM")

                    If Not dictMonth.Exists(key) Then
                        dictMonth.Add key, False
                    End If

                    ' ?? CORRECT COLUMN USED HERE
                    If Val(wsLoc.Cells(i, "Z").Value) > 0 Then
                        dictMonth(key) = True
                    End If

                End If
            End If
        End If
    Next i

    ' STEP 2: output only those months
    For i = 2 To lastRowLocal
        If Trim(wsLoc.Cells(i, "C").Value) = millerName Then
            If IsDate(wsLoc.Cells(i, "B").Value) Then
                If wsLoc.Cells(i, "B").Value < cutoffDate Then

                    key = Format(wsLoc.Cells(i, "B").Value, "YYYY-MM")

                    If dictMonth.Exists(key) Then
                        If dictMonth(key) = True Then

                            wsRes.Cells(outRow, 1).Value = wsLoc.Cells(i, "B").Value
                            wsRes.Cells(outRow, 2).Value = wsLoc.Cells(i, "D").Value
                            wsRes.Cells(outRow, 3).Value = wsLoc.Cells(i, "G").Value
                            wsRes.Cells(outRow, 4).Value = wsLoc.Cells(i, "M").Value
                            wsRes.Cells(outRow, 5).Value = wsLoc.Cells(i, "N").Value
                            wsRes.Cells(outRow, 6).Value = wsLoc.Cells(i, "Y").Value
                            wsRes.Cells(outRow, 7).Value = wsLoc.Cells(i, "Z").Value

                            outRow = outRow + 1
                            dataFound = True

                        End If
                    End If

                End If
            End If
        End If
    Next i

    ' =========================

    If dataFound Then
        wsRes.Range("A4:G" & outRow - 1).Sort Key1:=wsRes.Range("A4"), Order1:=xlAscending, Header:=xlNo

        Dim totalRow As Long
        totalRow = outRow + 1

        wsRes.Cells(totalRow, 2).Value = "TOTAL"
        wsRes.Cells(totalRow, 3).Formula = "=SUM(C4:C" & outRow - 1 & ")"
        wsRes.Cells(totalRow, 5).Formula = "=SUM(E4:E" & outRow - 1 & ")"
        wsRes.Cells(totalRow, 6).Formula = "=SUM(F4:F" & outRow - 1 & ")"
        wsRes.Cells(totalRow, 7).Formula = "=SUM(G4:G" & outRow - 1 & ")"

        wsRes.Range("A3:G" & totalRow).Font.Bold = True

        On Error Resume Next
        wsRes.ListObjects(1).Delete
        On Error GoTo 0

        wsRes.ListObjects.Add(xlSrcRange, wsRes.Range("A3:G" & totalRow), , xlYes).Name = "MillerTable"

        MsgBox "Data found for Miller: " & millerName, vbInformation
    Else
        wsRes.Range("A3:G3").Font.Bold = True
        MsgBox "No data found for Miller: " & millerName, vbExclamation
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
