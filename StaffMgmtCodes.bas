Attribute VB_Name = "StaffMgmtCodes"

Sub openStaffMgmt()
    Dim settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff

    SyncstaffData_FromGsheets
    
    With StaffMgmt.ListBox1
        
        .ColumnCount = 10
        
        .ColumnWidths = "50;60;60;100;120;85;50;40;1;50"
        
        .RowSource = "StaffData!" & Sheet6.Range("A2:J" & Sheet6.Range("A1").End(xlDown).row).Address
        
        .Selected(0) = True
    
    End With
    
    With StaffMgmt
        
        .Top = Application.Top + Application.Height / 2 - .Height / 2.45
        
        .Left = Application.Left + Application.Width / 2 - .Width / 2
        
        .Width = 650
    
    End With
    
    StaffMgmt.TabStrip1.Value = 0
    
    If StaffMgmt.Visible = False Then StaffMgmt.Show vbModeless
    
    settings.Restore
    
End Sub


Sub FilterPCO()

    Dim settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff

    Dim ws As Worksheet, resultsheet As Worksheet
    
    Dim a As Range
    
    Set ws = Sheet16
    
    'refreshAllContracts
    Set resultsheet = Sheet18
       
    Dim rng As Range
    Dim src As Range
    Dim startRow As Long
    Dim addr As String, addr2 As String
    Dim newTbl As ListObject
    Dim name As name
    
    resultsheet.Cells.Delete
    
    resultsheet.Range("A1").Value = "PCO"
    
    addr = Range("F1").Address
        
    On Error GoTo Handler
    
    For Each a In Sheet17.Range("PCOList")
    'For Listbox Source
    If a.Value <> "" Then
        
        resultsheet.Range("A2").Value = a.Value
        
        resultsheet.Range(addr).Resize(1, ws.ListObjects(1).ListColumns.count).Value = ws.Range("D17").Resize(1, ws.ListObjects(1).ListColumns.count).Value
    
        Set rng = ws.Range("D17:" & ws.Range("D17").End(xlToRight).Address).Resize(ws.ListObjects(1).ListRows.count + 1, ws.ListObjects(1).ListColumns.count)
        
        Set src = resultsheet.Range("A1:A2")
        
        Set drng = resultsheet.Range(addr).Resize(1, ws.ListObjects(1).ListColumns.count)
        
        On Error Resume Next
        
        rng.AdvancedFilter xlFilterCopy, src, drng
        
        On Error GoTo 0
        
        
        Set newTbl = resultsheet.ListObjects.Add(xlSrcRange, drng.CurrentRegion, , xlYes)
        
        newTbl.name = Replace(a.Value, " ", "")
        
        addr = resultsheet.Range(addr).End(xlToRight).Offset(0, 5).Address
    
    End If
    
    Next a
    
    settings.Restore
    
    Exit Sub
    
Handler:
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "FilterPCO"
    
    MsgBox "error occourred Check Update table for error information"

    settings.Restore
End Sub



Sub openPortfolioMgmt()

    Dim settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    'ThisWorkbook.Worksheets("MyContractsTable").ListObjects(1).QueryTable.Refresh BackgroundQuery:=False
    
    clearFilters Sheet16
    
    'applyAdvFilt Sheet16
    
    FilterPCO
    
    Dim addr As String
    
    addr = Sheet18.Range(Replace(Sheet17.Range("PCOList").Resize(1, 1).Value, " ", "")).Address
    
    With PortfolioMgmt.ListBox1
        
        .ColumnCount = 17
        
        .ColumnWidths = "40;60;120;70;70;150;60;70;70;30;50;50;1;50;30;30;2"
        
        .RowSource = "PCOprofiles!" & addr
        
        .Selected(0) = True
    
    End With
    
    'addr = Sheet18.Range(Replace(Sheet17.Range("PCOList").Resize(1, 1).Offset(1).Value, " ", "")).Address
    With PortfolioMgmt.ListBox2
        
        .ColumnCount = 17
        
        .ColumnWidths = "40;120;120;70;70;150;60;70;70;30;50;50;100;50;30;30;1"
        
        .RowSource = ""
        
        '.Selected(0) = True
    
    End With
    
    PortfolioMgmt.Field_5.Value = Sheet17.Range("PCOList").Resize(1, 1).Value
    
    PortfolioMgmt.Field_6.Value = ""
    
    If PortfolioMgmt.Visible = False Then PortfolioMgmt.Show
    
    
    With PortfolioMgmt
        
        .Top = Application.Top + Application.Height / 2 - .Height / 2.45
        
        .Left = Application.Left + Application.Width / 2 - .Width / 2
    
    End With
    
    settings.Restore
    
End Sub
