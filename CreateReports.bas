Attribute VB_Name = "CreateReports"
Option Explicit

Sub reportAdvFilt()

    Dim i               As Long
    Dim sh              As Worksheet
    Dim src As Range, drng As Range, rng As Range
    Dim lastcol As Long
    Dim parentsheet As Worksheet
    Dim rows As Long
    Dim col As Long
    Dim j As Long
    Dim dlfrFinder As Integer
    Dim sourceArr As Variant
    Dim DestinArr As Variant
    Dim resultcolCount As Long
    Dim k As Long
    
    ThisWorkbook.Activate
    
    ThisWorkbook.ActiveSheet.Range("D18").Select
    
        Set parentsheet = ThisWorkbook.ActiveSheet
    

    
    If Left(Sheet12.Range("Position"), 3) = "PCO" And parentsheet.name <> Sheet16.name And parentsheet.name <> Sheet19.name Then
        
        Set sh = Sheet14
        
        Sheet14.Range("DA1").Value = "PCO"
        
        Sheet14.Range("DA2").Value = Range("pName").Value
        
        Set rng = Sheet8.Range("A1").CurrentRegion
        
        Set src = Sheet14.Range("DA1").CurrentRegion
        
        Set drng = Sheet14.Range("A1").CurrentRegion

        sh.ListObjects(1).Resize rng
        
        rng.AdvancedFilter xlFilterCopy, src, drng
        
        sh.ListObjects(1).Resize drng.CurrentRegion
    
    Else
        
        Set sh = Sheet8
    
    End If
    
    lastcol = sh.Range("A1").End(xlToRight).Column
    
    ApplyFilters parentsheet
    
    'Timer.PrintTime "Apply Filter Conditions"
    
   
    Set rng = sh.Cells(1, db_primaryKey).CurrentRegion
    
    '
    
    Set src = sh.Cells(1, db_filters).CurrentRegion
    
    Set drng = sh.Cells(1, db_FilterResult).Resize(1, rng.Columns.count)
    
    drng.CurrentRegion.Offset(1, 0).Clear
    
    '
    Set src = sh.Range("DA1").CurrentRegion
    
    Set drng = sh.Range("BA1").Resize(1, rng.Columns.count)
    
    On Error Resume Next
    
    rng.AdvancedFilter xlFilterCopy, src, drng
    
    On Error GoTo 0
    
End Sub


Sub CreateReport()

    Dim Settings As New ExclClsSettings

    Dim reportWb As Workbook
    
    Dim reportTable As ListObject
    
    Dim ptTable As PivotTable
    
    Dim pCache As PivotCache
    
    Dim pRange As Range
    
    Dim LR As Long, LC As Long
    
    Dim dataSheet As Worksheet
    
    Dim ws As Worksheet
    
    Dim Yesno As String
    
    Dim pivotSheet As Worksheet
    
    Dim arr As Variant
    
    reportAdvFilt
    
    Settings.TurnOff
    
    If InStr(1, Sheet12.Range("Position").Value, "PCO", vbTextCompare) > 0 And ActiveSheet.name <> Sheet16.name Then
    
        Set dataSheet = Sheet14
    
    Else
    
        Set dataSheet = Sheet8
        
    End If
    
    If dataSheet.Range("BA1").CurrentRegion.rows.count = 1 Then
    
        Settings.Restore
        
        MsgBox ("Not enough data available to create a Report!")
        
        Exit Sub
    
    End If
            
    Set reportWb = Workbooks.Add
    
    arr = dataSheet.Range("BA1").CurrentRegion.Value
    
    Set ws = reportWb.Worksheets(1)
    
    ws.name = "Report Data"
    
    ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
    
    Set reportTable = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    
    reportTable.name = "Data"
    
    'copies table from the selected page and creates a new book
    
    Set pivotSheet = reportWb.Worksheets.Add(, reportWb.ActiveSheet)
    
    pivotSheet.name = "PRIORITY SUMMARY"
    
    
''''''''''''PRIORITY REPORT
    Set pRange = ws.Range("A1").CurrentRegion
    
    Set pCache = reportWb.PivotCaches.Create(xlDatabase, pRange)
    
    pivotSheet.Range("A1").Value = "PRIORITY REPORT"
    
    Set ptTable = pCache.CreatePivotTable(pivotSheet.Cells(5, 1), "Summary1")
    
    With pivotSheet.PivotTables("Summary1").PivotFields("Priority")
        
        .Orientation = xlRowField
        .position = 1
        
    End With
    
    
    With pivotSheet.PivotTables("Summary1").PivotFields("PCO")
    
        .Orientation = xlColumnField
        .position = 1
    
    End With
    
    With pivotSheet.PivotTables("Summary1").PivotFields("Primary_Key")
        
        .Orientation = xlDataField
        .position = 1
        
    End With
    
''''''''''''PCO PORTFOLIO OVERVIEW REPORT
    
    Set pivotSheet = reportWb.Worksheets.Add(, reportWb.ActiveSheet)
    
    pivotSheet.name = "PCO SUMMARY"
    
    pivotSheet.Range("A1").Value = "PCO PORTFOLIO OVERVIEW REPORT"

    
    Set ptTable = pCache.CreatePivotTable(pivotSheet.Cells(5, 1), "Summary2")
    
        With pivotSheet.PivotTables("Summary2").PivotFields("PCO")
        
        .Orientation = xlRowField
        .position = 1
        
    End With
    
    With pivotSheet.PivotTables("Summary2").PivotFields("Primary_Key")
        
        .Orientation = xlDataField
        .position = 1
        
    End With

''''''''''''Contract Types Report

    Set pivotSheet = reportWb.Worksheets.Add(, reportWb.ActiveSheet)
    
    pivotSheet.name = "CONTRACT TYPES SUMMARY"
    
    pivotSheet.Range("A1").Value = "CONTRACT TYPES REPORT"

    
    Set ptTable = pCache.CreatePivotTable(pivotSheet.Cells(5, 1), "Summary3")
    
        With pivotSheet.PivotTables("Summary3").PivotFields("Type")
        .Orientation = xlRowField
        .position = 1
        
    End With
    
    With pivotSheet.PivotTables("Summary3").PivotFields("Primary_Key")
        
        .Orientation = xlDataField
        .position = 1
        
    End With
''''''''''''Contract in Each term Report

    Set pivotSheet = reportWb.Worksheets.Add(, reportWb.ActiveSheet)
    
    pivotSheet.name = "CONTRACTS TERM SUMMARY"
    
    pivotSheet.Range("A1").Value = "CONTRACT IN EACH TERM REPORT"

    
    Set ptTable = pCache.CreatePivotTable(pivotSheet.Cells(5, 1), "Summary4")
    
        With pivotSheet.PivotTables("Summary4").PivotFields("Current Renewal Period")
        .Orientation = xlRowField
        .position = 1
        
    End With
    
    With pivotSheet.PivotTables("Summary4").PivotFields("Primary_Key")
        
        .Orientation = xlDataField
        .position = 1
        
    End With
Dim fileName As String
    
    ws.Activate
    
    Settings.Restore
    
    
    Yesno = MsgBox("Do you want to save this report workbook?", vbYesNo, "Save File?")
    
    If Yesno = vbNo Then
    
        Settings.Restore
        
        Exit Sub
        
    End If
    
    fileName = InputBox("Enter the name of the report")
    
    If fileName = "" Then fileName = "Report_" & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
    
    Dim obj As Object
    
    Set obj = CreateObject("scripting.FileSystemObject")
    
    If Not obj.FolderExists(Replace(ThisWorkbook.FullName, ThisWorkbook.name, "Portfolio Reports")) Then
        
        obj.CreateFolder (Replace(ThisWorkbook.FullName, ThisWorkbook.name, "Portfolio Reports"))
        
    End If
    
    reportWb.SaveAs Replace(ThisWorkbook.FullName, ThisWorkbook.name, "Portfolio Reports") & "\" & fileName
    

    
End Sub

Sub CreatePivotTable()
        Dim Settings As New ExclClsSettings
Settings.TurnOn


End Sub
