Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo SafeExit

If Target.Cells.CountLarge > 1 Then Exit Sub

Application.EnableEvents = False
Application.ScreenUpdating = False

Dim wsParty As Worksheet, wsMiller As Worksheet
Dim f As Range
Dim r As Long
Dim lastRowParty As Long, lastRowMiller As Long
Dim valG, valK, valM, valO, valP

' Using error checking to ensure sheets exist
On Error Resume Next
Set wsParty = Sheets("PARTY")
Set wsMiller = Sheets("MILLER")
On Error GoTo SafeExit

r = Target.Row

'---------------------------------
'PARTY LOOKUP (DICTIONARY BASED)
'---------------------------------
If Target.Column = 4 Then

    Dim dictParty As Object
    Dim i As Long
    Dim keyP As String
    Dim searchValP As String

    searchValP = Trim(CStr(Target.Value))

    If searchValP <> "" Then

        Set dictParty = CreateObject("Scripting.Dictionary")

        lastRowParty = wsParty.Cells(wsParty.Rows.Count, "A").End(xlUp).Row

        'Load dictionary
        For i = 2 To lastRowParty
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
            Me.Cells(r, "K").Value = dictParty(searchValP)(1)
        End If

    End If

    GoTo SafeExit
End If

'---------------------------------
'MILLER VALIDATION (DICTIONARY BASED)
'---------------------------------
If Target.Column = 3 Then

    Dim dict As Object
    Dim i2 As Long
    Dim key As String
    Dim searchVal As String

    searchVal = Trim(CStr(Target.Value))

    'Clear old values
    Me.Cells(r, "M").ClearContents
    Me.Cells(r, "C").Font.Color = vbBlack

    If searchVal <> "" Then

        Set dict = CreateObject("Scripting.Dictionary")

        lastRowMiller = wsMiller.Cells(wsMiller.Rows.Count, "B").End(xlUp).Row

        'Load data into dictionary
        For i2 = 2 To lastRowMiller
            key = Trim(CStr(wsMiller.Cells(i2, "B").Value))
            
            If key <> "" Then
                dict(key) = wsMiller.Cells(i2, "C").Value
            End If
        Next i2

        'Lookup
        If dict.Exists(searchVal) Then
            Me.Cells(r, "M").Value = dict(searchVal)
            Me.Cells(r, "C").Font.Color = vbBlack
        Else
            Me.Cells(r, "C").Font.Color = vbRed
        End If

    End If

    GoTo SafeExit
End If

'---------------------------------
'RATE COPY
'---------------------------------
If Target.Column = 9 Then
    Me.Cells(r, "J").Value = Target.Value
    ' Removed GoTo SafeExit to allow flow to PROFIT calculation
End If

'---------------------------------
'COMMISSION CALCULATION
'---------------------------------
If Target.Column = 7 Or Target.Column = 11 Or Target.Column = 13 Then

    valG = Me.Cells(r, "G").Value

    If valG = "" Then
        Me.Range("L" & r & ":O" & r & ",Q" & r).ClearContents
    Else
        valK = Me.Cells(r, "K").Value
        valM = Me.Cells(r, "M").Value

        Me.Cells(r, "L").Value = valG * valK
        Me.Cells(r, "N").Value = valG * valM
        Me.Cells(r, "O").Value = Me.Cells(r, "L").Value + Me.Cells(r, "N").Value
        ' Recalculate Q based on O and existing P
        Me.Cells(r, "Q").Value = Val(Me.Cells(r, "O").Value) - Val(Me.Cells(r, "P").Value)
    End If

    ' Only exit if the change was NOT in Column 7 (G)
    If Target.Column <> 7 Then GoTo SafeExit
End If

'---------------------------------
'CREDIT UPDATE / MANUAL P ENTRY (MODIFIED)
'---------------------------------
' This block handles both Column 16 (P) and any manual updates
If Target.Column = 16 Then

    valO = Me.Cells(r, "O").Value
    valP = Target.Value

    If valO <> "" Then
        ' Recalculate Q = O - P
        Me.Cells(r, "Q").Value = Val(valO) - Val(valP)
    End If

    GoTo SafeExit
End If

'---------------------------------
'PROFIT
'---------------------------------
' Listens to Column 7 (G), Column 9 (I), or Column 10 (J)
If Target.Column = 7 Or Target.Column = 9 Or Target.Column = 10 Then

    valG = Me.Cells(r, "G").Value

    If valG = "" Then
        Me.Cells(r, "W").ClearContents
    Else
        ' Recalculates W based on (J - I) * G
        Me.Cells(r, "W").Value = (Val(Me.Cells(r, "J").Value) - Val(Me.Cells(r, "I").Value)) * Val(valG)
    End If

End If

SafeExit:
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub
