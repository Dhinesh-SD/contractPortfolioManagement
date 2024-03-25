Attribute VB_Name = "WorksheetProtectionCode"
'This Page Contains All the code required to protect the worksheets in this workbook
Option Explicit
 Dim ws As Worksheet, a As Range

Sub protectWorksheet()

'Code to Protect All Worksheet

    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Range("A1").Value = "NavTo" Then
            
            protc ws
        
        End If
    
    Next ws
    
    protc Sheet6

End Sub

Sub unProtectWorksheet()
    'Code to unprotect All Worksheets
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Range("A1").Value = "NavTo" Then
            
            unprotc ws
        
        End If
    
    Next ws
    
    unprotc Sheet6
    
End Sub
Sub unProtectWorkbook(Optional booksA As Workbook)
'Code to unprotect any workbook passed as an argument
    Dim count As Long, maxsheets As Long
    
    count = 0
    
    maxsheets = booksA.Sheets.count
    
    For Each ws In booksA.Worksheets
        
        unprotc ws
        
        count = count + 1
    
    Next ws

End Sub

Sub ProtectWorkbook(Optional booksA As Workbook)
'Code to protect any workbook passed as an argument

    Dim count As Long, maxsheets As Long
    
    count = 0
    
    maxsheets = booksA.Sheets.count
    
    For Each ws In booksA.Worksheets
        
        protc ws
        
        count = count + 1
    
    Next ws

End Sub
Sub tempUnprotc()
 unprotc
End Sub


Sub unprotc(Optional Active As Worksheet)
'If not passed as an argument this code will unprotect any active sheet regardless of the workbook this code resides in
    Dim settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    If Active Is Nothing Then
    
    Set Active = ActiveSheet
    
    End If
    
    Active.Unprotect Password:=""
    
    settings.Restore

End Sub

Sub protc(Optional Active As Worksheet)
'If not passed as an argument this code will protect any active sheet regardless of the workbook this code resides in

    Dim settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    If Active Is Nothing Then
        
        Set Active = ActiveSheet
    
    End If
    
    'Password field is empty and if you need to enter a password change it here
    
    Active.Protect Password:="", DrawingObjects:=True, UserInterfaceonly:=True, Contents:=True, Scenarios:= _
            False, AllowSorting:=True, AllowFiltering:=True
    
    settings.Restore
    
End Sub



