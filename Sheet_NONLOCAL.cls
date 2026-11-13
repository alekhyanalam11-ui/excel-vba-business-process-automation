Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo exitHandler

If Target.Cells.CountLarge > 1 Then GoTo exitHandler

Application.EnableEvents = False

Dim wsParty As Worksheet
Dim wsMiller As Worksheet
Dim f As Range
Dim r As Long
Dim searchVal As String

' Check if sheets exist to prevent errors
On Error Resume Next
Set wsParty = ThisWorkbook.Sheets("PARTY")
Set wsMiller = ThisWorkbook.Sheets("MILLER")
On Error GoTo exitHandler

r = Target.Row

'========================
' PARTY LOOKUP (DICTIONARY BASED - FIXED)
'========================
If Not Intersect(Target, Me.Columns("D")) Is Nothing Then

    Dim dictParty As Object
    Dim lastRow As Long
    Dim i As Long
    Dim keyP As String
    Dim searchValP As String

    searchValP = Trim(CStr(Target.Value))

    If searchValP <> "" Then

        Set dictParty = CreateObject("Scripting.Dictionary")

        lastRow = wsParty.Cells(wsParty.Rows.Count, "A").End(xlUp).Row

        'Build dictionary
        For i = 1 To lastRow
            keyP = Trim(CStr(wsParty.Cells(i, "A").Value))

            If keyP <> "" Then
                dictParty(keyP) = Array( _
                    wsParty.Cells(i, "B").Value, _
                    wsParty.Cells(i, "C").Value)
            End If
        Next i

        'Lookup
        If dictParty.Exists(searchValP) Then

            Me.Cells(r, "D").Value = dictParty(searchValP)(0)

            'Always update commission rate whenever party changes
            Me.Cells(r, "K").Value = dictParty(searchValP)(1)
            
            'Recalculate Party Commission
            Me.Cells(r, "L").Value = Me.Cells(r, "G").Value * Me.Cells(r, "K").Value

            'Update Balance Commission
            Me.Cells(r, "N").Value = Val(Me.Cells(r, "L").Value) - Val(Me.Cells(r, "M").Value)


            Me.Cells(r, "D").Font.Color = vbBlack

        Else
            'Keep your exact validation logic
            If Not IsNumeric(searchValP) Then
                Me.Cells(r, "D").Font.Color = vbRed
            Else
                Me.Cells(r, "D").Font.Color = vbBlack
            End If
        End If

    End If

End If

'========================
' MILLER LOOKUP (DICTIONARY - NO CONFLICT)
'========================
If Not Intersect(Target, Me.Columns("C")) Is Nothing Then

    Dim dictMillerLocal As Object
    Dim iM As Long
    Dim lastRowMillerLocal As Long
    Dim keyM As String
    Dim searchValLocal As String

    searchValLocal = Trim(CStr(Target.Value))

    'Clear old values
    Me.Cells(r, "Q").ClearContents
    Me.Cells(r, "C").Font.Color = vbBlack

    If searchValLocal <> "" Then

        Set dictMillerLocal = CreateObject("Scripting.Dictionary")

        lastRowMillerLocal = wsMiller.Cells(wsMiller.Rows.Count, "B").End(xlUp).Row

        'Build dictionary
        For iM = 2 To lastRowMillerLocal
            keyM = Trim(CStr(wsMiller.Cells(iM, "B").Value))

            If keyM <> "" Then
                dictMillerLocal(keyM) = wsMiller.Cells(iM, "C").Value
            End If
        Next iM

        'Lookup
        If dictMillerLocal.Exists(searchValLocal) Then

            Me.Cells(r, "C").Font.Color = vbBlack

            If Me.Cells(r, "Q").Value = "" Then
                Me.Cells(r, "Q").Value = dictMillerLocal(searchValLocal)
            End If

        Else
            Me.Cells(r, "C").Font.Color = vbRed
        End If

    End If

End If

'========================
' RATE COPY (UNCHANGED)
'========================
If Not Intersect(Target, Me.Columns("I")) Is Nothing Then
    Me.Cells(r, "J").Value = Me.Cells(r, "I").Value
End If

'========================
' COMMISSIONS (RECALC N = L - M ADDED)
'========================
If Not Intersect(Target, Me.Range("G:G,K:K,Q:Q")) Is Nothing Then

    Me.Cells(r, "L").Value = Me.Cells(r, "G").Value * Me.Cells(r, "K").Value
    Me.Cells(r, "R").Value = Me.Cells(r, "G").Value * Me.Cells(r, "Q").Value

    ' Added Val() to handle empty cells during subtraction
    Me.Cells(r, "N").Value = Val(Me.Cells(r, "L").Value) - Val(Me.Cells(r, "M").Value)
    Me.Cells(r, "T").Value = Val(Me.Cells(r, "R").Value) - Val(Me.Cells(r, "S").Value)

End If

'========================
' CREDIT UPDATE (PARTY) - MANUAL COLUMN M ENTRY
'========================
If Not Intersect(Target, Me.Columns("M")) Is Nothing Then
    ' Recalculate N = L - M automatically on manual entry
    Me.Cells(r, "N").Value = Val(Me.Cells(r, "L").Value) - Val(Me.Cells(r, "M").Value)
End If

'========================
' UPDATED M CREDIT LOGIC (COL S)
'========================
If Not Intersect(Target, Me.Columns("S")) Is Nothing Then

    If Me.Cells(r, "S").Value = "" Then
        Me.Cells(r, "T").Value = Me.Cells(r, "R").Value
    Else
        Me.Cells(r, "T").Value = Val(Me.Cells(r, "R").Value) - Val(Me.Cells(r, "S").Value)
    End If

End If

'========================
' PROFIT (UNCHANGED)
'========================
If Not Intersect(Target, Me.Range("G:G,I:I,J:J")) Is Nothing Then
    Me.Cells(r, "Z").Value = (Me.Cells(r, "J").Value - Me.Cells(r, "I").Value) * Me.Cells(r, "G").Value
End If

exitHandler:
Application.EnableEvents = True

End Sub
