VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Password 
   Caption         =   "Password"
   ClientHeight    =   3516
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8136
   OleObjectBlob   =   "Password.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub Btn_Login_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_Login

End Sub

Private Sub Btn_Login_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_Login

End Sub

Private Sub Btn_Login_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_Login

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

Private Sub Field_1_Change()
    
    setProperties Field_1, ThemeUf.FilledTB

End Sub



Private Sub Btn_Login_Click()
Dim i As Long
    
    i = CInt(Me.Label2.Caption)
    
    
    If SHA1(Field_1.Value) = Sheet6.Cells(i, 9).Value Then
        
        SignIn.Label3.Caption = "Correct Password"
        
        Unload Me
    
    Else
         
         Me.Field_1.BorderColor = &HFF&
         
         Dim yesNo As String
         
         yesNo = MsgBox("Incorrect Password!" & vbNewLine & "Would you like to try again?", vbYesNo, "Try Again?")
         
        If yesNo = vbYes Then
            
            Me.Field_1.Value = ""
        
        Else
            
            SignIn.Label3.Caption = "Wrong Password"
            
            Unload Me
        
        End If
    
    End If
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



Private Sub Field_1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 13 Then
    
        Btn_Login_Click
    
    End If

End Sub

Private Sub UserForm_Initialize()

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
    
    Me.Field_1.PasswordChar = "*"
    
    setProperties Field_1, ThemeUf.FilledTB
    
    Me.BackColor = ThemeUf.PageBackGround.BackColor

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_Login
    
    Exit_highlight Btn_ShowHide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If SignIn.Label3.Caption <> "Correct Password" Then
        
        SignIn.Label3.Caption = "Wrong Password"
    
    End If

End Sub
