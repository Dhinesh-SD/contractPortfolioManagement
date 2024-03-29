VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SignIn 
   Caption         =   "SELECT PROFILE"
   ClientHeight    =   3696
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7692
   OleObjectBlob   =   "SignIn.frx":0000
End
Attribute VB_Name = "SignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_SignIn_Click()
    
   ' Dim Timer As New 'TimerCls
    Dim Settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    'Timer.start
    
    Settings.TurnOn
        
    Settings.TurnOff
    
    'Timer.PrintTime "Turn-Off Settings"
    
    SyncstaffData_FromGsheets "Mandatory"
    
    Dim i As Long
    Dim lrow As Long
    Dim Email As String
    Dim position As String
    Dim Security As String
    Dim fileLocation As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet, ws As Worksheet
    Dim HomePage As Worksheet
    Dim myTable As ListObject
    Dim contractsTableSheet As Worksheet
    Dim src As Range, drng As Range
    Dim notFound As Boolean
    Dim r As Long
    Dim j As Long
    
    unProtectWorksheet
    
    notFound = True
            
    'Timer.PrintTime "Update Staff Data Table"
    
    
    
    Set HomePage = Sheet1
    
    Set ws1 = Sheet6
    
    Set contractsTableSheet = Sheet14
    
    Email = LCase(Me.Field_1.Value)
    
    If Right(Email, 13) <> "@nebraska.gov" Then
    
        Email = Email & "@nebraska.gov"
    
    End If
    
    'On Error GoTo Handler
    
    lrow = ws1.Range("A1").End(xlDown).row
    
    For i = 2 To lrow
        
        If LCase(ws1.Cells(i, 5).Value) = Email Then
            
            position = ws1.Cells(i, 6).Value
            
            Security = ws1.Cells(i, 8).Value
            
            If ws1.Cells(i, 9).Value = "" Then
                
                With NewPassword
                    
                    .Field_2.Caption = ws1.Cells(i, 4).Value
                    
                    .Top = Application.Top + Application.Height / 2 - .Height / 2
                    
                    .Left = Application.Left + Application.Width / 2 - .Width / 2
                    
                    .Label2.Caption = i
                    
                    .Field_1.PasswordChar = "*"
                
                End With
            
                NewPassword.Show
                
                Unload Me
                
                SignIn.Field_1.Value = Email
                
                If ws1.Cells(i, 9).Value <> "" Then SignIn.Btn_SignIn.Value = True
                
                protc
                
                Exit Sub
            
            End If
            
            With Password
                
                .Top = Application.Top + Application.Height / 2 - .Height / 2
                
                .Left = Application.Left + Application.Width / 2 - .Width / 2
                
                .Label2.Caption = i
                
                .Field_1.PasswordChar = "*"
            
            End With
            
            'Timer.PrintTime "Show Password Userform"
            
            Password.Show
            
            'Timer.start
            
            If LCase(Me.Label3.Caption) = "wrong password" Then
            
                Settings.Restore
                
                Exit Sub
                
            End If
            
            'Check if Profile is already logged in if yes exit sub
            
            If ws1.Cells(i, "J").Value = "Logged-In" Then
                                
                'Timer.PrintTime " LoginFailed"
                
                MsgBox "Login Failed:Profile Already in Use!", , "Login Error"
                
                Unload Me
                
                protc
                
                Settings.Restore
                
                Exit Sub
            
            End If
            
            r = 2
            
            For j = 1 To 8
            
                Sheet12.Cells(r, 1).Value = ws1.Cells(1, j).Value
                Sheet12.Cells(r, 2).Value = ws1.Cells(i, j).Value
                r = r + 1
            Next j
                
                   
            HomePage.Shapes("Info_profileName").Visible = msoCTrue
            'Change functionality of app based on Position of user
            unprotc ws1
            
                ws1.Cells(i, "L").Value = ""
                
                ws1.Cells(i, "J").Value = "Logged_In"
                
                SyncStaffData_ToGsheets
                
            protc ws1
            
            updateLog ThisWorkbook, "Logged In", "Sign in Procedure"
            
            'Timer.PrintTime " Login Successful and Update Values"
            
            For Each ws In ThisWorkbook.Worksheets
                
                If ws.Range("A1").Value = "Nav_To" And ws.name <> Sheet1.name Then
                
                    unprotc ws
                    
                    ws.Range("A1").Value = "NavTo"
                    
                End If
            Next ws
            
                            
            If Left(position, 3) <> "PCO" And position <> "Administrator" Then
                
                'display My contracts page
                
                ThisWorkbook.Worksheets("My Contracts").Range("A1").Value = "Nav_To"
                
            End If

            AddShapes Sheet1
            
            ''Timer.PrintTime " AddNew Shapes in Homepage"
            
            notFound = False
            
            Exit For
            
        End If
    
    Next i
    
    If notFound = True Then
        
        MsgBox "Login Failed:Profile Not found!", , "Login Error"
    
        Settings.Restore
        
        'Timer.PrintTime " LoginFailed"
        
        Exit Sub
        
    End If
    
    Dim rng As Range, srcRng As Range
    
    renameShapes
    
    Sheet14.Range("A1").CurrentRegion.Offset(1).Clear
    
    Sheet14.Range("DA1").CurrentRegion.Clear
    
    Sheet14.Range("DA1").Value = "PCO"
     
     
     If Left(Sheet12.Range("Position").Value, 3) = "PCO" Then
     
        Sheet14.Range("DA2").Value = Sheet12.Range("B5").Value
        Set rng = Sheet8.Range("A1").CurrentRegion
        Set srcRng = Sheet14.Range("DA1").CurrentRegion
        Set drng = Sheet14.Range("A1").CurrentRegion
        
        rng.AdvancedFilter xlFilterCopy, srcRng, drng
        
        Sheet14.ListObjects(1).Resize Range("A1").CurrentRegion
     End If
        
    protectWorksheet
    
    'Timer.PrintTime "Table Update"
    Dim sh As Worksheet
    
    For Each sh In ThisWorkbook.Worksheets
        
        If sh.Range("A1").Value = "NavTo" Then
            
            sh.Range("A3").Value = True
        
        End If
        
    Next sh
    
    'Timer.PrintTime " Change Page ""A1"" Values "
    
    Unload Me
    
    Settings.Restore
    
    'Timer.PrintTime "Restore Settings"
    
    Exit Sub
    
Handler:

    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff
    
    emergencySignOut
    
        updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "SignIn Btn_SignIn_Click: Unsuccessful"
            
        Unload Me
    
    Settings.Restore

End Sub


Private Sub Btn_SignIn_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_SignIn

End Sub

Private Sub Btn_SignIn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_SignIn

End Sub

Private Sub Btn_SignIn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_SignIn

End Sub

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
        
        
        If InStr(1, cntrl.name, "Field", vbTextCompare) > 0 And InStr(1, cntrl.name, "Field_Header", vbTextCompare) = 0 Then
        
            baseField cntrl
        
        End If
    Next cntrl
    
    With Me
    
        .Top = Application.Top + Application.Height / 2 - .Height / 2
        
        .Left = Application.Left + Application.Width / 2 - .Width / 2
    
    End With
    
    Me.BackColor = ThemeUf.PageBackGround.BackColor
    Me.ForeColor = ThemeUf.PageBackGround.ForeColor

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_SignIn

End Sub
