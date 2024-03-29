VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StaffMgmt 
   Caption         =   "Staff Management"
   ClientHeight    =   11220
   ClientLeft      =   60
   ClientTop       =   288
   ClientWidth     =   12780
   OleObjectBlob   =   "StaffMgmt.frx":0000
End
Attribute VB_Name = "StaffMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_AddUser_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_AddUser

End Sub

Private Sub Btn_AddUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_AddUser
    Exit_highlight Btn_EditUser

End Sub

Private Sub Btn_AddUser_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_AddUser
    
End Sub


Private Sub Btn_AddUser_Click()
    Me.TabStrip1.Value = 1
End Sub

Private Sub Btn_EditUser_Click()
    Me.TabStrip1.Value = 0
End Sub

Private Sub Btn_EditUser_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_EditUser

End Sub

Private Sub Btn_EditUser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_EditUser
    Exit_highlight Btn_AddUser

End Sub

Private Sub Btn_EditUser_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_EditUser
    
End Sub


Private Sub Btn_UpdateInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_UpdateInfo

End Sub

Private Sub Btn_UpdateInfo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_UpdateInfo
    Exit_highlight Btn_DeleteStaff
End Sub

Private Sub Btn_UpdateInfo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_UpdateInfo
    
End Sub


Private Sub Btn_UpdateInfo_Click()

    Dim Settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff
    
    Dim wb As Workbook, wb1 As Workbook
    
    Dim fileLocation1 As String, fileLocation As String
    
    Dim i As Long, row As Long
    
    Dim sh As Worksheet
    
    Dim rng As Range
    
    'Email check
    Dim ws As Worksheet, ws1 As Worksheet, ws2 As Worksheet
    
    If (Len(Me.Field_5.Value) > 13 And LCase(Right(Me.Field_5.Value, 13)) <> "@nebraska.gov") Or Len(Me.Field_5.Value) < 13 Then
        
        MsgBox ("Enter a valid work email Id!")
        
        Exit Sub
    End If
    
    Set ws = Sheet6
    
    For Each rng In ws.Range(ws.ListObjects(1).name & "[Staff_ID]")
    
        If rng.Value = Me.Field_1.Caption Then
        
            row = rng.row
            
            Exit For
            
        End If
    
    Next rng
    
    If ws.Cells(row, "J").Value = "Logged_In" Then
        
        MsgBox ("Profile Currently in use! Edit after User Logs Out!")
        
        Settings.Restore
        
        Exit Sub
    
    End If
    

    'ws.Cells(row, 1).Value = Me.Field_1.Caption
    
    ws.Cells(row, 2).Value = Me.Field_2.Value
    
    ws.Cells(row, 3).Value = Me.Field_3.Value
    
    If ws.Cells(row, 4).Value <> Me.Field_4.Caption And Left(ws.Cells(row, "F").Value, 3) = "PCO" Then
    
        
        For Each rng In Sheet8.Range(Sheet8.ListObjects(1).name & "[PCO]")
        
            If rng.Value = ws.Cells(row, 4) Then
            
                rng.Value = Me.Field_4.Caption
                
                Sheet8.Cells(rng.row, "AT").Value = ""
            
            End If
            
        Next rng
    
        syncData
        
        For Each sh In ThisWorkbook.Worksheets
        
        If sh.Range("A1").Value = "NavTo" Then
            
            sh.Range("A3").Value = True
        
        End If
    
    Next sh
    
    End If
    
    ws.Cells(row, 4).Value = Me.Field_4.Caption
    
    ws.Cells(row, 5).Value = Me.Field_5.Value
    
    ws.Cells(row, 6).Value = Me.Field_6.Caption
    
    ws.Cells(row, 7).Value = Me.Field_7.Value
    
    ws.Cells(row, 8).Value = Me.Field_8.Value
    
    ws.Cells(row, "K").Value = "No"
    
    ws.Cells(row, "L").Value = ""
    
    SyncStaffData_ToGsheets
    
    Dim k As Long
    
    k = 4
    
    'Updating The list of PCO's for comboBox lists
    
    For i = 2 To Sheet6.Range("A1").End(xlDown).row
        
        If Left(Sheet6.Range("F" & i).Value, 3) = "PCO" Then
            
            Sheet17.Range("G" & k).Value = Sheet6.Range("D" & i).Value
            
            k = k + 1
        
        End If
    
    Next i
    
    Me.changesMade.Caption = True
    
    MsgBox "Changes Saved"
    
    Settings.Restore
    
    Exit Sub
    
Handler:

    Settings.Restore
    
    updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "Staff Mgmt Btn_AddNewStaff_Click(Delete Staff): Unsuccessful"
    
    
End Sub

Private Sub Btn_AddNewStaff_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_AddNewStaff

End Sub

Private Sub Btn_AddNewStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_AddNewStaff

End Sub

Private Sub Btn_AddNewStaff_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_AddNewStaff
    
End Sub

Private Sub Btn_AddNewStaff_Click()
    Dim Settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff
    
    Dim fileLocation As String
    Dim i As Long
    Dim ws As Worksheet
    'Email Check
    
    If (Len(Me.Field_5.Value) > 13 And LCase(Right(Me.Field_5.Value, 13)) <> "@nebraska.gov") Or Len(Me.Field_5.Value) < 13 Then
        
        MsgBox ("Enter a valid work email Id!")
        
        Exit Sub
    
    End If
   
    On Error GoTo Handler
    
    Me.ListBox1.RowSource = ""
    
    Set ws = Sheet6
    
    unprotc ws
    
        i = ws.ListObjects(1).ListRows.count + 2
    
        ws.Cells(i, 1).Value = i - 1
        
        ws.Cells(i, 2).Value = Me.Field_2.Value
        
        ws.Cells(i, 3).Value = Me.Field_3.Value
        
        ws.Cells(i, 4).Value = Me.Field_4.Caption
        
        ws.Cells(i, 5).Value = Me.Field_5.Value
        
        ws.Cells(i, 6).Value = Me.Field_6.Caption
        
        ws.Cells(i, 7).Value = Me.Field_7.Value
        
        ws.Cells(i, 8).Value = Me.Field_8.Value
        
        ws.Cells(i, 10).Value = "Loged_Out"
        
        ws.Cells(i, 11).Value = "No"
        
        ws.Cells(i, 12).Value = ""
    
    protc ws
    
    SyncStaffData_ToGsheets

    Set ws = Nothing
    
    Me.ListBox1.RowSource = "StaffData!" & Sheet6.Range("A2:H" & Sheet6.Range("H1").End(xlDown).row).Address


    Dim sh As Worksheet
    
    For Each sh In ThisWorkbook.Worksheets
        
        If sh.Range("A1").Value = "NavTo" Then
            
            sh.Range("A3").Value = True
        
        End If
    
    Next sh
    
    Settings.Restore
    
    Me.TabStrip1.Value = 0
    
    MsgBox "New User Added"
    
    Me.changesMade.Caption = True
    
    Exit Sub

Handler:
    
    
    updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "Staff Mgmt Btn_AddNewStaff_Click(Add New Staff): Unsuccessful"

    Settings.Restore

End Sub


Private Sub Btn_DeleteStaff_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_DeleteStaff

End Sub

Private Sub Btn_DeleteStaff_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_DeleteStaff
    Exit_highlight Btn_UpdateInfo

End Sub

Private Sub Btn_DeleteStaff_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_DeleteStaff
    
End Sub

Private Sub Btn_DeleteStaff_Click()

    Dim Settings As New ExclClsSettings
    
    Settings.TurnOn
    
    Settings.TurnOff
    
    
    Dim wb As Workbook, wb1 As Workbook
    
    Dim fileLocation1 As String, fileLocation As String
    
    Dim i As Long, row As Long
    
    Dim rng As Range
    
    Dim SelectedField_1 As Long
    
    Dim ws As Worksheet, sh As Worksheet
    
    Dim Yesno As String
    
    Yesno = MsgBox("Do You Want to Delete " & Me.Field_4.Caption & "'s Profile?", vbYesNo, "Confirmation")
    
    If Yesno = vbNo Then
    
        Settings.Restore
        
        Exit Sub
    
    End If
    
    SelectedField_1 = CInt(Me.Field_1.Caption)
    
    SyncstaffData_FromGsheets
    
    Set ws = Sheet6
    
   
       
    For Each rng In ws.Range(ws.ListObjects(1).name & "[Staff_ID]")
    
        If rng.Value = SelectedField_1 Then
        
            row = rng.row
            
        End If
    
    Next rng
    
    If ws.Cells(row, "J").Value = "Logged_In" Then
                
        MsgBox ("Profile Currently Logged_In! Cannot Delete Until User Logs Out!")
        
        Settings.Restore
        
        Exit Sub
    
    End If

 Me.ListBox1.RowSource = ""
unprotc ws
        
    ws.Cells(row, "K").Value = "Yes"
    
    ws.Cells(row, "L").Value = ""
    
    If Left(ws.Cells(row, "F").Value, 3) = "PCO" Then
    
        For Each rng In Sheet8.Range("User_Data[PCO]")
        
            If rng.Value = ws.Cells(row, "D").Value Then
            
                rng.Value = ws.Cells(row, "F").Value
                
                Sheet8.Cells(rng.row, "AT").Value = ""
                
                
            End If
        
        Next rng
        
        syncData
        
        For Each sh In ThisWorkbook.Worksheets
        
            If sh.Range("A1").Value = "NavTo" Then
                
                sh.Range("A3").Value = True
            
            End If
        
        Next sh
    
    End If

    SyncStaffData_ToGsheets
    
    protc ws

    Me.changesMade.Caption = True
    
    Unload Me
    
    openStaffMgmt
            

Settings.Restore

Exit Sub

Handler:

Settings.Restore

updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "Staff Mgmt Btn_AddNewStaff_Click(Delete Staff): Unsuccessful"


End Sub





Private Sub ActBtn_ResetPassword_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick ActBtn_ResetPassword

End Sub

Private Sub ActBtn_ResetPassword_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight ActBtn_ResetPassword

End Sub

Private Sub ActBtn_ResetPassword_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight ActBtn_ResetPassword
    
End Sub

Private Sub ActBtn_ResetPassword_Click()

    Dim Settings As New ExclClsSettings
    
    Settings.TurnOn
    
    Settings.TurnOff
    
    Dim wb As Workbook, wb1 As Workbook
    
    Dim fileLocation1 As String, fileLocation As String
    
    Dim i As Long, row As Long
    
    Dim ws As Worksheet, sh As Worksheet
    
    Dim rng As Range
    
    'SyncstaffData_FromGsheets
    
    Set ws = Sheet6
       
    For Each rng In ws.Range(Sheet6.ListObjects(1).name & "[Staff_ID]")
    
        If rng.Value = CInt(Me.Field_1.Caption) Then
        
            row = rng.row
            
        End If
    
    Next rng
    
    
    'Profile in use check
    
    If ws.Cells(row, "J").Value = "Logged-In" Then
               
        MsgBox ("Profile Currently in use! Cannot reset password Until User Logs Out!")
        
        Settings.Restore
        
        Exit Sub
    
    End If

    On Error GoTo Handler
        
        unprotc ws
        
            ws.Cells(row, "I").Value = ""
            
            ws.Cells(row, "L").Value = ""
            
        protc ws
        
        SyncStaffData_ToGsheets
        
        MsgBox "Password Successfuly Reset!"
        
    Settings.Restore
        
    Exit Sub
    
Handler:
    
    updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "Staff Mgmt Btn_ResetPassword_Click(Reset Password Staff): Unsuccessful"
        
    Settings.Restore

End Sub




Private Sub ListBox1_Click()
Dim i As Long
    
    For i = 0 To Me.ListBox1.ListCount - 1
    
        If Me.ListBox1.Selected(i) = True And Me.TabStrip1.Value = 0 Then
        
            Me.Field_1.Caption = Me.ListBox1.List(i, 0)
            
            Me.Field_2.Value = Me.ListBox1.List(i, 1)
            
            Me.Field_3.Value = Me.ListBox1.List(i, 2)
            
            Me.Field_4.Caption = Me.ListBox1.List(i, 3)
            
            Me.Field_5.Value = Me.ListBox1.List(i, 4)
            
            Me.Field_6.Caption = Me.ListBox1.List(i, 5)
            
            Me.Field_7.Value = Me.ListBox1.List(i, 6)
            
            Me.Field_8.Value = Me.ListBox1.List(i, 7)
        
        End If
    
    Next i

End Sub



Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_EditUser
    Exit_highlight Btn_AddUser
    
End Sub

Private Sub Mini_Page_Heading_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Exit_highlight Btn_AddUser
End Sub

Private Sub TabStrip1_Change()
    Dim i As Long
    
    If Me.TabStrip1.Value = 0 Then
        
        Me.Mini_Page_Heading.Caption = vbNewLine & UCase("Edit Existing Staff Information")
        
        For i = 0 To Me.ListBox1.ListCount - 1
            
            If Me.ListBox1.Selected(i) = True And Me.TabStrip1.Value = 0 Then
                
                Me.Field_1.Caption = Me.ListBox1.List(i, 0)
                
                Me.Field_2.Value = Me.ListBox1.List(i, 1)
                
                Me.Field_3.Value = Me.ListBox1.List(i, 2)
                
                
                Me.Field_4.Caption = Me.ListBox1.List(i, 3)
                
                Me.Field_5.Value = Me.ListBox1.List(i, 4)
                
                Me.Field_6.Caption = Me.ListBox1.List(i, 5)
                
                Me.Field_7.Value = Me.ListBox1.List(i, 6)
                
                Me.Field_8.Value = Me.ListBox1.List(i, 7)
            
            End If
        
        Next i
        
        Me.Btn_UpdateInfo.Visible = True
        
        Me.Btn_AddNewStaff.Visible = False
        
        Me.Btn_DeleteStaff.Visible = True
        
        highlight Me.Btn_EditUser
        
        Exit_highlight Me.Btn_AddUser
        
    Else
        
        Me.Mini_Page_Heading.Caption = vbNewLine & UCase("Add New Staff Member")
        
        Me.Field_1.Caption = Me.ListBox1.ListCount + 1
        
        Me.Field_6.Caption = "Guest_" & CInt(Me.ListBox1.ListCount) - 9
        
        Me.Field_4.Caption = Me.Field_6.Caption
        
        Me.Field_2.Value = ""
        
        Me.Field_3.Value = ""
        
        Me.Field_5.Value = ""
        
        Me.Field_7.Value = ""
        
        Me.Field_8.Value = "User"
        
        Me.Btn_UpdateInfo.Visible = False
        
        Me.Btn_AddNewStaff.Visible = True
        
        Me.Btn_DeleteStaff.Visible = False
        
        Exit_highlight Me.Btn_EditUser
        
        highlight Me.Btn_AddUser
    
    End If

End Sub



Private Sub Field_2_Change()

If Me.Field_2.Value = "" And Me.Field_3.Value = "" Then
    Me.Field_4.Caption = Me.Field_6.Caption
    Me.Field_5.Value = LCase(Me.Field_6.Caption & "@nebraska.gov")
Else
    Me.Field_4.Caption = UCase(Me.Field_2.Value & " " & Me.Field_3.Value)
    Me.Field_5.Value = LCase(Me.Field_2.Value & "." & Me.Field_3.Value & "@nebraska.gov")

End If
End Sub

Private Sub Field_3_Change()
If Me.Field_2.Value = "" And Me.Field_3.Value = "" Then
    Me.Field_4.Caption = Me.Field_6.Caption
    Me.Field_5.Value = Me.Field_6.Caption & "@nebraska.gov"

Else
    Me.Field_4.Caption = UCase(Me.Field_2.Value & " " & Me.Field_3.Value)
    Me.Field_5.Value = LCase(Me.Field_2.Value & "." & Me.Field_3.Value & "@nebraska.gov")

End If
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
    
    Me.Mini_Page_Heading.Caption = vbNewLine & UCase("Edit Existing Staff Information")
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim cntrl As Control
    
    For Each cntrl In Me.Controls
    
        If Left(cntrl.name, 3) = "Btn" Then Exit_highlight cntrl
    
        If Left(cntrl.name, 3) = "Act" Then highlight cntrl
    Next cntrl

End Sub

Private Sub UserForm_Terminate()

    If Me.changesMade.Caption = "True" Then
    
        ThisWorkbook.Save
        
    End If
    
End Sub
