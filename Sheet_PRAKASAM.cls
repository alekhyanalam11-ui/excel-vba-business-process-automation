Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.CountLarge > 1 Then Exit Sub

    'K1 - PRAKASAM PENDING DELIVERY
    If Not Intersect(Target, Me.Range("K1")) Is Nothing Then

        Extract_Prakasham_Pending_Delivery

    'L1 - PRAKASAM PENDING COMMISSION
    ElseIf Not Intersect(Target, Me.Range("L1")) Is Nothing Then

        Extract_Prakasham_Pending_Commission

    End If

End Sub
