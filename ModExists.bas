Attribute VB_Name = "ModExists"
 Function Exists(strCtrlName As String, CtrlIndex As Integer) As Boolean
    
    '  Determines if a control exists
    Dim ctrl As Control
    Exists = False

     For Each ctrl In Screen.ActiveForm.Controls

        If ctrl.Name Like strCtrlName Then
            If ctrl.Index = CtrlIndex Then
                Exists = True
                Exit Function  ' Quit function once form has been found.
            End If
        End If
    Next

End Function
