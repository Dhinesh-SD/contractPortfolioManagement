VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewPassword 
   Caption         =   "Set New Password"
   ClientHeight    =   3732
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   7320
   OleObjectBlob   =   "NewPassword.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub setProperties(cntrl1 As Control, cntrl2 As Control)
'
'    cntrl1.BackColor = cntrl2.BackColor
'    cntrl1.ForeColor = cntrl2.ForeColor
'    cntrl1.BackStyle = cntrl2.BackStyle
'    cntrl1.BorderStyle = cntrl2.BorderStyle
'    cntrl1.BorderColor = cntrl2.BorderColor
'
'End Sub
'
'Private Sub highlight(cntrl As Control)
'
'If Not cntrl.BackColor = ThemeUf.SampleActive.BackColor Then
'
'    setProperties cntrl, ThemeUf.SampleActive
'
'End If
'
'End Sub
'
'
'Private Sub Exit_highlight(cntrl As Control)
'
'If Not cntrl.BackColor = ThemeUf.SampleInactive.BackColor Then
'
'    setProperties cntrl, ThemeUf.SampleInactive
'
'End If
'
'End Sub
'
'Private Sub Onclick(cntrl As Control)
'
'    setProperties cntrl, ThemeUf.onClick_Btn
'
'End Sub


Private Sub Btn_CreatePassword_Click()

    Dim ws As Worksheet
    
    Set ws = Sheet6
    
    unprotc ws
        
        ws.Cells(CInt(Me.Label2.Caption), 9).Value = SHA1(Me.Field_1.Value)
        
        ws.Cells(CInt(Me.Label2.Caption), 12).Value = ""
    
    protc ws
    
    SyncStaffData_ToGsheets
        
    Unload Me

End Sub


Private Sub Btn_CreatePassword_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_CreatePassword

End Sub

Private Sub Btn_CreatePassword_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_CreatePassword
    
End Sub

Private Sub Btn_CreatePassword_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Exit_highlight Btn_CreatePassword

End Sub

Private Sub Btn_ShowHide_Click()

If Btn_ShowHide.Caption = "Hide Password" Then
        
        Me.Field_1.PasswordChar = "*"
        
        Btn_ShowHide.Caption = "Show Password"
        
    Else
        
        Me.Field_1.PasswordChar = ""
        
        Btn_ShowHide.Caption = "Hide Password"
    
    End If
End Sub



Private Sub Btn_ShowHide_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Onclick Btn_ShowHide
End Sub

Private Sub Btn_ShowHide_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    highlight Btn_ShowHide
End Sub

Private Sub Btn_ShowHide_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Exit_highlight Btn_ShowHide
End Sub

Private Sub UserForm_Initialize()

    
    'frm.Visible = False
    
    Dim cntrls As Control
    
    For Each cntrls In Me.Controls
    
        If InStr(1, cntrls.name, "Field", vbTextCompare) > 1 And InStr(1, cntrls.name, "Field_Header") = 0 Then
        
            setProperties cntrls, ThemeUf.FilledTB
        
        End If
        
        If Left(cntrls.name, 3) = "Btn" Then
            
            Exit_highlight cntrls
            
        End If
        
        If cntrls.name = "Heading" Then
        
            setProperties cntrls, ThemeUf.SampleInactive
        End If
        
        If InStr(1, cntrls.name, "Field_Header") > 0 Then

            setProperties cntrls, ThemeUf.PageBackGround
            
        End If
        
        
    Next cntrls


End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Exit_highlight Btn_ShowHide
    
    Exit_highlight Btn_CreatePassword

End Sub
