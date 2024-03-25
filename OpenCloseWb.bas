Attribute VB_Name = "OpenCloseWb"
Option Explicit

Function OpenWb(filelocation As String, Optional rdOnly As Boolean = False) As Workbook
Dim wb As Workbook
If filelocation = "" Then
    Set OpenWb = Nothing
    Exit Function
End If
Set wb = Workbooks.Open(filelocation, , rdOnly)
    'Check if workbook is open if yes return nothing else return workbook
    If wb.ReadOnly = True And rdOnly = False Then
        wb.Close False
        MsgBox "Database Busy/ Opened by another user!"
        Set OpenWb = Nothing
        Exit Function
    End If
    
    Set OpenWb = wb
    
End Function
