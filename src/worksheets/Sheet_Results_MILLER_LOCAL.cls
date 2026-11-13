Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' --- ALWAYS KEEP Q1 FIXED ---
    ' This part runs every time you move your selection,
    ' ensuring Q1 is ALWAYS bold and has the text.
    Application.EnableEvents = False
    With Me.Range("Q1")
        .Value = "CLICK HERE FOR SINGLE ENTRY"
        .Font.Bold = True
    End With
    Application.EnableEvents = True

    ' --- TRIGGER CALCULATION ONLY ON Q1 CLICK ---
    If Not Intersect(Target, Me.Range("Q1")) Is Nothing Then
        
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        Dim lastRow As Long
        Dim i As Long
        
        ' 1. CLEANUP: Delete any row where Column B says "TOTAL"
        ' This is the "Magic Fix" that stops the numbers from doubling.
        lastRow = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
        For i = lastRow To 2 Step -1
            If Trim(UCase(Me.Cells(i, "B").Value)) = "TOTAL" Then
                Me.Rows(i).Delete
            End If
        Next i

        ' 2. Find new last row of data in Column G
        lastRow = Me.Cells(Me.Rows.Count, "G").End(xlUp).Row
        
        ' 3. Delete rows where Column G (M DEBIT) is 0 or Empty
        For i = lastRow To 2 Step -1
            If Me.Cells(i, "G").Value = 0 Or IsEmpty(Me.Cells(i, "G")) Then
                Me.Rows(i).Delete
            End If
        Next i
        
        ' 4. Recalculate Totals on a fresh clean row
        lastRow = Me.Cells(Me.Rows.Count, "G").End(xlUp).Row
        
        Me.Cells(lastRow + 1, "B").Value = "TOTAL"
        Me.Cells(lastRow + 1, "B").Font.Bold = True
        
        ' Summing Columns C, E, and G
        Me.Cells(lastRow + 1, "C").Value = Application.WorksheetFunction.Sum(Me.Range("C2:C" & lastRow))
        Me.Cells(lastRow + 1, "E").Value = Application.WorksheetFunction.Sum(Me.Range("E2:E" & lastRow))
        Me.Cells(lastRow + 1, "F").Value = Application.WorksheetFunction.Sum(Me.Range("F2:F" & lastRow))
        Me.Cells(lastRow + 1, "G").Value = Application.WorksheetFunction.Sum(Me.Range("G2:G" & lastRow))
        
        Me.Range("K1").Font.Bold = True
        
        MsgBox "DATA FILTERED & TOTALS UPDATED", vbInformation, "Success"
        
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
    End If
End Sub
