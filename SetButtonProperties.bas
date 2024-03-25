Attribute VB_Name = "SetButtonProperties"
Public Sub setProperties(cntrl1 As Control, cntrl2 As Control)

    cntrl1.BackColor = cntrl2.BackColor
    cntrl1.BackStyle = cntrl2.BackStyle
    cntrl1.ForeColor = cntrl2.ForeColor
    cntrl1.BackStyle = cntrl2.BackStyle
    If TypeName(cntrl1) <> "CommandButton" Then
        cntrl1.BorderStyle = cntrl2.BorderStyle
        cntrl1.BorderColor = cntrl2.BorderColor
    End If
    
End Sub

Public Sub highlight(cntrl As Control)

If Not cntrl.BackColor = ThemeUf.SampleActive.BackColor Then

    setProperties cntrl, ThemeUf.SampleActive
    
End If

End Sub


Public Sub Exit_highlight(cntrl As Control)

If Not cntrl.BackColor = ThemeUf.SampleInactive.BackColor Then

    setProperties cntrl, ThemeUf.SampleInactive
    
End If

End Sub

Public Sub EmptyField(cntrl As Control)

    setProperties cntrl, ThemeUf.EmptyTB

End Sub

Public Sub Onclick(cntrl As Control)

    setProperties cntrl, ThemeUf.onClick_Btn

End Sub

Public Sub baseField(cntrl As Control)

    setProperties cntrl, ThemeUf.FilledTB

End Sub
