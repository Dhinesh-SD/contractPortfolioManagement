VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Status 
   Caption         =   "Status"
   ClientHeight    =   2184
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7932
   OleObjectBlob   =   "Status.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
Dim cntrl As Control
    
    For Each cntrl In Me.Controls
           
        If Left(cntrl.name, 3) = "Btn" Then Exit_highlight cntrl
        
        If Right(cntrl.name, 7) = "heading" Then
        
            setProperties cntrl, ThemeUf.sampleHeading
            
        End If
        
        
        If InStr(1, cntrl.name, "Field_Header", vbTextCompare) > 0 Then
        
            cntrl.ForeColor = ThemeUf.PageBackGround.ForeColor
        
        End If
        
        
    Next cntrl
    
    With Me
    
        .Top = Application.Top + Application.Height / 2 - .Height / 2
        
        .Left = Application.Left + Application.Width / 2 - .Width / 2
    
    End With
    
    Me.BackColor = ThemeUf.PageBackGround.BackColor
    Me.ForeColor = ThemeUf.PageBackGround.ForeColor

End Sub
