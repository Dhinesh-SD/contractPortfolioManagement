Attribute VB_Name = "FilesAndFolderBrowse"
Option Explicit

Sub CreateProfile()

'The purpose of the VBA code is to automate the process of creating user folders and subfolders based on data in a worksheet. Here's a summary of what the code does:

'It creates a folder for each user listed in the worksheet, using the data in column 4 to generate the folder name.
'For each user folder, it creates subfolders based on the data in column 10, separated by commas.
'If a user folder or subfolder already exists, it skips that folder and continues to the next one.
'For each user, it checks if a workbook with the same name as the user's folder already exists in the user's folder.
'If the workbook does not exist, it creates a new workbook from a template and saves it in the user's folder.
'If the workbook already exists, it opens the workbook and updates the data in the worksheet with the data from the corresponding row in the worksheet.

'Procedure to Create User folders and sub folders
Application.screenUpdating = False
    Dim staffSheet As Worksheet
    Dim i As Long
    Dim staffLastRow As Long
    Dim folderpath As String, folderName As String
    
    Set staffSheet = Sheet6
    
    Dim objFSO As Object, objFolder As Object, objFile As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    staffLastRow = staffSheet.Range("A1").End(xlDown).row
    
    Dim userdataFolder As String
    
    Dim subFolders As Collection
    Dim subfold() As String, foldName As Variant
    Dim Filename As String
    Dim wb As Workbook
    Dim j As Long
    For i = 2 To staffLastRow
        ReDim subfold(1 To 15)
        subfold = Split(staffSheet.Cells(i, 10).Value, ",")
        folderpath = staffSheet.Cells(i, 9).Value
        folderName = UCase(Replace(Replace(staffSheet.Cells(i, 4).Value, " ", "_"), ".", ""))
        'Check if userfolder Exists If not create folder
        If Not (objFSO.FolderExists(folderpath & "\" & folderName)) Then
            objFSO.CreateFolder (folderpath & "\" & folderName)
            For Each foldName In subfold
                If Not (objFSO.FolderExists(folderpath & "\" & folderName & "\" & foldName)) And foldName <> "" Then
                    objFSO.CreateFolder (folderpath & "\" & folderName & "\" & foldName)
                End If
            Next foldName
        End If
        userdataFolder = Replace(folderpath, "Profile Data", "")
        Filename = staffSheet.Cells(i, 6).Value & ".xlsx"
        'If Workbook not exists, create New Workbook from template
        If Dir(userdataFolder & Filename) <> "" Then
            
        Else
            Set wb = Workbooks.Add(userdataFolder & "Template.xlsx")
            wb.SaveAs userdataFolder & Filename
            For j = 1 To staffSheet.Cells(1, 1).End(xlToRight).Column
                wb.Worksheets("Sheet3").Range("B" & j + 1).Value = staffSheet.Cells(i, j).Value
            Next j
              wb.Worksheets("Sheet3").Range("B13").Value = "logged-off"
            wb.Close True
        End If
    Next
Application.screenUpdating = True

End Sub

