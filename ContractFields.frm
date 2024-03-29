VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContractFields 
   Caption         =   "Contract Field Management"
   ClientHeight    =   7224
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11988
   OleObjectBlob   =   "ContractFields.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContractFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Btn_UpdateInfo_Click()

    Me.ListBox1.RowSource = ""
    
    
    
    unprotc Sheet7
    
    Sheet7.Range("D" & CInt(Me.Field_1.Caption) + 1).Value = Me.Field_4.Value
    
    Sheet7.Range("H" & CInt(Me.Field_1.Caption) + 1).Value = ""

    syncFieldAccess
    
    Me.ListBox1.RowSource = "'" & Sheet7.name & "'!" & Sheet7.ListObjects(1).DataBodyRange.Resize(Sheet7.ListObjects(1).ListRows.count, 6).Address
       
    protc Sheet7

    

End Sub


Private Sub Field_Header_6_Click()

End Sub

Private Sub ListBox1_Click()
Dim i As Long

With Me.ListBox1

For i = 0 To .ListCount - 1
    
    If .Selected(i) = True Then
        
        Me.Field_1.Caption = .List(i, 0)
        
        Me.Field_2.Caption = .List(i, 1)
        
        Me.Field_4.Value = CBool(.List(i, 3))
    
        Exit For
        
    End If

Next i

End With
End Sub

Private Sub UserForm_Initialize()
Dim cntrl As Control
    
    For Each cntrl In Me.Controls
    
        If Left(cntrl.name, 2) = "Bg" Then setProperties cntrl, ThemeUf.PageBackGround
        
        If Right(cntrl.name, 5) = "title" Then
            
            setProperties cntrl, ThemeUf.SampleInactive
                    
        End If
        
        
        If Left(cntrl.name, 3) = "Act" Then highlight cntrl
        
        If Left(cntrl.name, 3) = "Btn" Then Exit_highlight cntrl
        
        If LCase(Right(cntrl.name, 7)) = "heading" Then
        
            setProperties cntrl, ThemeUf.sampleHeading
            
        End If
        
        
        If InStr(1, cntrl.name, "Field_Header", vbTextCompare) > 0 Then
        
            cntrl.ForeColor = ThemeUf.PageBackGround.ForeColor
        
        End If
        
        If TypeName(cntrl) = "ComboBox" Or TypeName(cntrl) = "TextBox" Then
        
            cntrl.FontName = "Arial"
            cntrl.Font.Size = 10
        
        End If
        
        If InStr(1, cntrl.name, "Field", vbTextCompare) > 0 And InStr(1, cntrl.name, "Field_Header", vbTextCompare) = 0 Then
        
            baseField cntrl
        
        End If
    Next cntrl
    
    Me.BackColor = ThemeUf.PageBackGround.BackColor
    
    Me.ForeColor = ThemeUf.PageBackGround.ForeColor
    
    
    With Me.ListBox1
        
        .ColumnCount = 6
        
        .ColumnWidths = "20;120;0;120;180;80"
        
        .RowSource = "'" & Sheet7.name & "'!" & Sheet7.ListObjects(1).DataBodyRange.Resize(Sheet7.ListObjects(1).ListRows.count, 6).Address
    
    End With

        Me.ListBox1.Selected(0) = True


    With Me
    
        .Top = Application.Top + Application.Height / 2 - Me.Height / 2
        .Left = Application.Left + Application.Width / 2 - Me.Width / 2
        
    End With
    
End Sub
