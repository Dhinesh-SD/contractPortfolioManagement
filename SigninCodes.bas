Attribute VB_Name = "SigninCodes"
Option Explicit
Sub SignInCode()
    Dim Settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff
    
    SyncstaffData_FromGsheets
    
    Dim i As Long
    Dim staffSheet As Worksheet
    
    SignIn.Show
    
    Theme
    
    Settings.Restore
    
End Sub


Function GetIPAddress()
    Const strComputer As String = "."   ' Computer name. Dot means local computer
    Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
    Dim strIPAddress As String

    ' Connect to the WMI service
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    ' Get all TCP/IP-enabled network adapters
    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' Get all IP addresses associated with these adapters
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = strIPAddress & Join(IPAddress, ", ")
        End If
    Next

    GetIPAddress = strIPAddress
End Function

Sub test()

Debug.Print GetIPAddress

End Sub

Sub SignOutCode()

    Dim Settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff
    
    Dim rng As Range
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim Yesno As String
    
    Yesno = MsgBox("Would you like to SignOut?", vbYesNo, "Sign-Out Confirmation")
    
    If Yesno = vbNo Then Exit Sub
    
    unProtectWorksheet
    
    
    Set ws = Sheet6

            unprotc ws

            For Each rng In Sheet6.Range(Sheet6.ListObjects(1).name & "[Staff_ID]")
                
                If rng.Value = Sheet12.Range("B2").Value Then
                
                    ws.Cells(rng.row, "L").Value = ""
                    
                    ws.Cells(rng.row, "J").Value = "Logged_Out"
                
                End If
            
            Next rng
            
            updateLog ThisWorkbook, "Logged Out", "Sign out Procedure"
            
            protc ws
            
            SyncStaffData_ToGsheets
            
            For Each ws In ThisWorkbook.Worksheets
                
                If ws.Range("A1").Value = "NavTo" And ws.name <> Sheet1.name Then ws.Range("A1").Value = "Nav_To"
            
            Next ws
            
            Sheet16.Range("A1").Value = "NavTo"
            Sheet12.Range("A2:B11").ClearContents
            Sheet1.Shapes("Info_ProfileName").TextFrame.Characters.Text = ""
     
    
    protectWorksheet
    
    ThisWorkbook.Save
    
    Settings.Restore
    
    AddShapes Sheet1
    
    Theme

End Sub

Sub emergencySignOut()
    
    Dim rng As Range
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim Yesno As String
    
   
    unProtectWorksheet
    
    If Sheet12.Range("B2").Value = "" Then

        Exit Sub
    
    End If
    
Set ws = Sheet6

            unprotc ws

            For Each rng In Sheet6.Range(Sheet6.ListObjects(1).name & "[Staff_ID]")
                
                If rng.Value = Sheet12.Range("B2").Value Then
                
                    ws.Cells(rng.row, "L").Value = ""
                    
                    ws.Cells(rng.row, "J").Value = "Logged_Out"
                
                End If
            
            Next rng
            
            
            
            updateLog ThisWorkbook, "Logged Out", "Emergency Sign out Procedure"
            
            protc ws
            
            SyncStaffData_ToGsheets
            
            
            For Each ws In ThisWorkbook.Worksheets
                
                If ws.Range("A1").Value = "NavTo" And ws.name <> Sheet1.name Then ws.Range("A1").Value = "Nav_To"
            
            Next ws
            
            Sheet16.Range("A1").Value = "NavTo"
            Sheet12.Range("A2:B11").ClearContents
            
            Sheet1.Shapes("Info_ProfileName").TextFrame.Characters.Text = ""
     
    
    AddShapes Sheet1
    
    Theme
    
    protectWorksheet
    
    ThisWorkbook.Save



End Sub

Sub newt()
Sheet16.Range("A1").Value = "NavTo"
End Sub
