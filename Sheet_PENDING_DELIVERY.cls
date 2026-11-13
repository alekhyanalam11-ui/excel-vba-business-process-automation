Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.CountLarge > 1 Then Exit Sub

    On Error GoTo SafeExit

    Select Case Target.Address(False, False)

        Case "K1"
            UpdatePendingLocal
            UpdateAutomationHeader Target

        Case "K3"
            UpdatePendingNonLocal
            UpdateAutomationHeader Target

        Case "L1"
            Extract_PRASAD
            UpdateAutomationHeader Target

        Case "L3"
            Extract_MADHU
            UpdateAutomationHeader Target

        Case "L5"
            Extract_JEMILI
            UpdateAutomationHeader Target

        Case "L7"
            Extract_KALI
            UpdateAutomationHeader Target

        Case "L9"
            Extract_MK
            UpdateAutomationHeader Target

        Case "L11"
            Extract_ONG_VRR
            UpdateAutomationHeader Target

    End Select

SafeExit:

End Sub


Private Sub UpdateAutomationHeader(ByVal Target As Range)

    With Me.Range("D1")

        .Value = CStr(Target.Value) & _
                 "    Dt: " & Format(Date, "dd/mm/yyyy")

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .Font.Bold = True
        .Font.Size = 14

    End With

End Sub
