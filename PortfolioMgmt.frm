VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PortfolioMgmt 
   Caption         =   "Portfolio Management"
   ClientHeight    =   10044
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   19788
   OleObjectBlob   =   "PortfolioMgmt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PortfolioMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_Close_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_Close

End Sub

Private Sub Btn_Close_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_Close

End Sub

Private Sub Btn_Close_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_Close

End Sub



Private Sub Btn_RevertChanges_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_RevertChanges

End Sub

Private Sub Btn_RevertChanges_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_RevertChanges

End Sub

Private Sub Btn_RevertChanges_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_RevertChanges

End Sub


Private Sub Btn_SaveChanges_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_SaveChanges

End Sub

Private Sub Btn_SaveChanges_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_SaveChanges

End Sub

Private Sub Btn_SaveChanges_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_SaveChanges

End Sub


Private Sub Btn_MovetoP1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_MovetoP1

End Sub

Private Sub Btn_MovetoP1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_MovetoP1

End Sub

Private Sub Btn_MovetoP1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_MovetoP1

End Sub



Private Sub Btn_MovetoP2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_MovetoP2

End Sub

Private Sub Btn_MovetoP2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_MovetoP2

End Sub

Private Sub Btn_MovetoP2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_MovetoP2

End Sub



Private Sub Field_5_Change()
    Dim addr As String
    
    If Me.Field_5.Value = Me.Field_6.Value Or Me.Field_5.Value = "" Then
        
        Me.Field_5.Value = Me.Profile1_Heading.Caption
        
        Exit Sub
    
    ElseIf Me.Field_5.Value = "" Then
        
        Me.Profile1_Heading.Caption = Me.Field_5.Value
        
        Me.ListBox1.RowSource = ""
    
    Else
        
        Me.Profile1_Heading.Caption = Me.Field_5.Value
        
        FilterPCO
        
        addr = Range(Sheet18.ListObjects(Replace(Me.Profile1_Heading.Caption, " ", "")).name).Address
        
        Me.ListBox1.RowSource = "PCOprofiles!" & addr
    
    End If

End Sub

Private Sub Field_6_Change()

    Dim addr As String
    
    If Me.Field_5.Value = Me.Field_6.Value Or Me.Field_6.Value = "" Then
        
        Me.Field_6.Value = Me.Profile2_Heading.Caption
        
        Exit Sub
    
    ElseIf Me.Field_6.Value = "" Then
        
        Me.Profile2_Heading.Caption = Me.Field_5.Value
        
        Me.ListBox2.RowSource = ""
    
    Else
        
        Me.Profile2_Heading.Caption = Me.Field_6.Value
        
        FilterPCO
        
        addr = Range(Sheet18.ListObjects(Replace(Me.Profile2_Heading.Caption, " ", "")).name).Address
        
        Me.ListBox2.RowSource = "PCOprofiles!" & addr
    
    End If

End Sub

Private Sub Btn_MovetoP1_Click()

    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    Dim i As Long
    Dim primeKey As String
    Dim profile1 As String, profile2 As String
    Dim j As Long
    Dim searchCol As Long
    
    profile1 = Me.Profile1_Heading.Caption
    
    profile2 = Me.Profile2_Heading.Caption
    
    If profile1 = "" Or profile2 = "" Then
        
        MsgBox ("Select Valid Profiles!")
        
        Exit Sub
    
    End If
    
    For i = Me.ListBox2.ListCount - 1 To 0 Step -1
        
        If Me.ListBox2.Selected(i) = True Then
            
            primeKey = Me.ListBox2.List(i, 16)
            
            searchCol = 1
            
            For j = 2 To Sheet8.Cells(1, 1).End(xlDown).row
                
                If Sheet8.Cells(j, searchCol).Value = primeKey Then
                    
                    Sheet8.Cells(j, 2).Value = profile1
                    
                    Sheet8.Cells(j, "AT").Value = ""
                    
                    Exit For
                
                End If
            
            Next j
        
        End If
    
    Next i
    
    Me.ListBox1.RowSource = ""
    
    Me.ListBox2.RowSource = ""
        
        
    Sheet16.Range("D13:S15").ClearContents
    
    
    applyAdvFilt Sheet16
    
    FilterPCO
    
    Dim addr As String
    
    addr = Sheet18.Range(Replace(profile1, " ", "")).Address
    
    With Me.ListBox1
        
        .RowSource = "PCOprofiles!" & addr
    
    End With
    
    addr = Sheet18.Range(Replace(profile2, " ", "")).Address
    
    With Me.ListBox2
        
        .RowSource = "PCOprofiles!" & addr
    
    End With
    
    settings.Restore

End Sub

Private Sub Btn_MovetoP2_Click()
    Dim i As Long
    Dim primeKey As String
    Dim profile1 As String, profile2 As String
    Dim j As Long
    Dim searchCol As Long
    
    profile1 = Me.Profile1_Heading.Caption
    
    profile2 = Me.Profile2_Heading.Caption
    
    If profile1 = "" Or profile2 = "" Then
        
        MsgBox ("Select Valid Profiles!")
        
        Exit Sub
    
    End If
    
    For i = Me.ListBox1.ListCount - 1 To 0 Step -1
        
        If Me.ListBox1.Selected(i) = True Then
            
            primeKey = Me.ListBox1.List(i, 16)
            
            searchCol = 1
            
            For j = 2 To Sheet8.Cells(1, searchCol).End(xlDown).row
                
                If Sheet8.Cells(j, searchCol).Value = primeKey Then
                    
                    Sheet8.Cells(j, 2).Value = profile2
                    
                    Sheet8.Cells(j, "AT").Value = ""
                    
                    Exit For
                
                End If
            
            Next j
        
        End If
    
    Next i
    
    Me.ListBox1.RowSource = ""
    
    Me.ListBox2.RowSource = ""
    
        
        Sheet16.Range("D13:T15").ClearContents
    
    
    applyAdvFilt Sheet16
    
    FilterPCO
    
    
    Dim addr As String
    
    addr = Sheet18.Range(Replace(profile1, " ", "")).Address
    
    With Me.ListBox1
        
        .RowSource = "PCOprofiles!" & addr
    
    End With
    
        
    addr = Sheet18.Range(Replace(profile2, " ", "")).Address
    
    With Me.ListBox2
        
        .RowSource = "PCOprofiles!" & addr
    
    End With


End Sub

Private Sub Btn_SaveChanges_Click()
    
    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    Dim wb As Workbook
    Dim ws As Worksheet, ws2 As Worksheet
    
    On Error GoTo Handler:
    
    'folderLocation = Sheet1.Shapes("Info_Root_Dir").TextFrame.Characters.Text & "\User Data\"
    
    syncData
            
    On Error GoTo Handler:
    
    updateLog ThisWorkbook, "Updated PCO profiles by" & ThisWorkbook.Worksheets("Profile Information").Range("pName").Value, _
        "SaveChanges From Protfolio Management Tool"
                
    Unload Me
        
    MsgBox "Changes Updated!"
    
    settings.Restore
    
    Exit Sub
    
Handler:
    
    settings.TurnOn
    
    settings.TurnOff
    
    updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "SaveChanges From Protfolio Management: Unsuccessful"
    
    settings.Restore


End Sub

Private Sub Btn_RevertChanges_Click()

refreshAllContracts

openPortfolioMgmt

End Sub

Private Sub Btn_Close_Click()

Unload Me

'refreshAllContracts

End Sub

Private Sub Field_3_Change()

Dim i As Long
Dim j As Long



j = CInt(Me.SearchField1.Caption)

For i = 0 To Me.ListBox1.ListCount - 1

    Me.ListBox1.Selected(i) = False


Next i


For i = 0 To Me.ListBox1.ListCount - 1
    
    If InStr(1, Me.ListBox1.List(i, j), Me.Field_3.Value, vbTextCompare) > 0 And Me.Field_3.Value <> "" And Me.ListBox1.Selected(i) = False Then
        
        Me.ListBox1.Selected(i) = True
            
    End If
   
Next i

End Sub

Private Sub Field_4_Change()

Dim i As Long

Dim j As Long

j = CInt(Me.SearchField2.Caption)


For i = 0 To Me.ListBox2.ListCount - 1

    Me.ListBox2.Selected(i) = False


Next i

For i = 0 To Me.ListBox2.ListCount - 1
    
    If InStr(1, Me.ListBox2.List(i, j), Me.Field_4.Value, vbTextCompare) > 0 And Me.Field_4.Value <> "" And Me.ListBox2.Selected(i) = False Then
    
        Me.ListBox2.Selected(i) = True
        
    End If
    
Next i

End Sub



Private Sub Field_7_Change()

Dim j As Long


For j = 6 To Sheet18.Range("F1").End(xlToRight).Column

    If Sheet18.Cells(1, j).Value = Me.Field_7.Value Then
                
        Me.SearchField1.Caption = j - 6
        
        Exit For
    
    Else
    
        Me.SearchField1.Caption = 5
        
    End If
    
Next j



End Sub

Private Sub Field_8_Change()

Dim j As Long


For j = 6 To Sheet18.Range("F1").End(xlToRight).Column

    If Sheet18.Cells(1, j).Value = Me.Field_8.Value Then
                
        Me.SearchField2.Caption = j - 6
        
        Exit For
    
    Else
    
        Me.SearchField1.Caption = 5
        
    End If
    
Next j



End Sub

Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Exit_highlight Btn_MovetoP2
    
End Sub

Private Sub ListBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Exit_highlight Btn_MovetoP1
    
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
        
        If Right(cntrl.name, 7) = "heading" Then
        
            setProperties cntrl, ThemeUf.sampleHeading
            
        End If
        
        
        If InStr(1, cntrl.name, "Field_Header", vbTextCompare) > 0 Then
        
            cntrl.ForeColor = ThemeUf.PageBackGround.ForeColor
        
        End If
        
        If TypeName(cntrl) = "ComboBox" Or TypeName(cntrl) = "TextBox" Then
        
            cntrl.FontName = Arial
            cntrl.Font.Size = 10
        
        End If
        
        If InStr(1, cntrl.name, "Field", vbTextCompare) > 0 And InStr(1, cntrl.name, "Field_Header", vbTextCompare) = 0 Then
        
            baseField cntrl
        
        End If
    Next cntrl
    
    With Me
    
        .Top = Application.Top + Application.Height / 2 - .Height / 2
        
        .Width = Application.Left + Application.Width / 2 - .Width / 2
            
        .Field_7.RowSource = "'" & Sheet17.name & "'!$AM$4:$AM$20"
        
        .Field_8.RowSource = "'" & Sheet17.name & "'!$AM$4:$AM$20"
    
    End With
    

    
    Me.BackColor = ThemeUf.PageBackGround.BackColor
    Me.ForeColor = ThemeUf.PageBackGround.ForeColor
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_Close
    Exit_highlight Btn_MovetoP1
    Exit_highlight Btn_MovetoP2
    Exit_highlight Btn_RevertChanges
    Exit_highlight Btn_SaveChanges

End Sub
