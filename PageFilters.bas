Attribute VB_Name = "PageFilters"
Option Explicit
' Column placeholders for database
'Public Enum DbColumns

    Public Const db_primaryKey = 1
    Public Const db_PCO = db_primaryKey + 1
    Public Const db_Type = db_PCO + 1
    Public Const db_contrNumber = db_Type + 1
    Public Const db_CLMSNum = db_contrNumber + 1
    Public Const db_DocType = db_CLMSNum + 1
    Public Const db_Rfx = db_DocType + 1
    Public Const db_Description = db_Rfx + 1
    Public Const db_Division = db_Description + 1
    Public Const db_DivContact = db_Division + 1
    Public Const db_TempPco1 = db_DivContact + 1
    Public Const db_TempPco2 = db_TempPco1 + 1
    Public Const db_Status = db_TempPco2 + 1
    Public Const db_Amd1 = db_Status + 1
    Public Const db_Amd2 = db_Amd1 + 1
    Public Const db_Amd3 = db_Amd2 + 1
    Public Const db_DeviType = db_Amd3 + 1
    Public Const db_DeviReason = db_DeviType + 1
    Public Const db_Agency = db_DeviReason + 1
    Public Const db_AgencyContact = db_Agency + 1
    Public Const db_Supplier = db_AgencyContact + 1
    Public Const db_DeviDate = db_Supplier + 1
    Public Const db_startDate = db_DeviDate + 1
    Public Const db_EndDate = db_startDate + 1
    Public Const db_NoOfRens = db_EndDate + 1
    Public Const db_EachRenDur = db_NoOfRens + 1
    Public Const db_MaxendDate = db_EachRenDur + 1
    Public Const db_Remarks = db_MaxendDate + 1
    Public Const db_Notes = db_Remarks + 1
    Public Const db_ExtensionDur = db_Notes + 1
    Public Const db_EstSpend = db_ExtensionDur + 1
    Public Const db_ContrLinked = db_EstSpend + 1
    Public Const db_Files = db_ContrLinked + 1
    Public Const db_Priority = db_Files + 1
    Public Const db_NextRenDate = db_Priority + 1
    Public Const db_DaysLeftFrRen = db_NextRenDate + 1
    Public Const db_CurrRenPeriod = db_DaysLeftFrRen + 1
    Public Const db_CLMSReqNum = db_CurrRenPeriod + 1
    Public Const db_SupplierABNum = db_CLMSReqNum + 1
    Public Const db_SharePointNum = db_SupplierABNum + 1
    Public Const db_DateEntered = db_SharePointNum + 1
    Public Const db_ContractDateEntered = db_DateEntered + 1
    Public Const db_RenewContract = db_ContractDateEntered + 1
    
    'Database filter result placeholder
    
    Public Const db_FilterResult = 53
    
    'Database filters
   Public Const db_filters = 105
    
    Public Const pg_FilterRow = 13

'End Enum


Sub applyFilterValidations()
'Use only if you edit fields in any of the pages
Dim i As Long
Dim j As Long
Dim filterRow As Long
Dim col As Long
Dim sh As Worksheet
Dim ws As Worksheet


Set sh = Sheet13
If ws Is Nothing Then Set ws = ActiveSheet
unprotc ws

col = ws.Range("D17").End(xlToRight).Column
'add Data Validations for filter Application for the respective cells based on the data type
filterRow = 13
For i = 4 To col
    If ws.Cells(filterRow + 4, i).Value <> "" Then
        For j = 4 To sh.Cells(3, 1).End(xlDown).row
            ''
            If Trim(ws.Cells(filterRow + 4, i).Value) = Trim(sh.Cells(j, 1).Value) Then
                With ws.Cells(filterRow, i).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="=" & sh.Cells(j, 2).Value
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With
                    ws.Cells(filterRow, i).Locked = False
                    ws.Cells(filterRow, i).Locked = False
                    ws.Cells(filterRow + 1, i).Locked = False
                    ws.Cells(filterRow + 2, i).Locked = False
                Exit For
            End If
        Next j
    End If
Next i
protc ws
End Sub


Private Sub applyTableHeaderValidations()
'Use only if you edit fields in any of the pages
Dim i As Long
Dim j As Long
Dim filterRow As Long
Dim col As Long
Dim sh As Worksheet
Dim ws As Worksheet


Set sh = Sheet13

If ws Is Nothing Then Set ws = ActiveSheet
unprotc ws

col = ws.Range("D17").End(xlToRight).Column
'add Data Validations for filter Application for the respective cells based on the data type
filterRow = 13
For i = 4 To col
            ''
            If ws.Cells(filterRow + 4, i).Value <> "Primary_Key" And ws.Cells(filterRow + 4, i).Value <> "Priority" And ws.Cells(filterRow + 4, i).Value <> "Type" Then
                With ws.Cells(filterRow + 4, i).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="='" & sh.name & "'!" & sh.Range("A4:A" & sh.Cells(4, 1).End(xlDown).row).Address
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With
                'ws.Cells(filterRow, i).Locked = False
            End If
Next i

ws.Range("D13:T17").Locked = False
protc ws
End Sub

Private Sub AddValidations()
Dim sh As Worksheet
Dim i As Long
Set sh = ActiveSheet
Dim tbl As ListObject



    If sh.name <> Sheet1.name And sh.Range("A1").Value = "NavTo" Then
        For i = 4 To sh.Cells(17, 4).End(xlToRight).Column
            For Each tbl In Sheet17.ListObjects
                On Error Resume Next
                With sh.Cells(14, i).Validation
                    .Delete
                    
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="=" & sh.Cells(17, i).Value & "List"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = False
                End With
            Next tbl

        Next i
    End If



End Sub

Sub printFilters(ws As Worksheet, resultsheet As Worksheet)
    
    Dim rng As Range
    Dim filters() As String
    Dim filterValsArr() As String
    Dim filtervals As String
    Dim filtHeaders As String
    Dim i As Long, j As Long
    Dim totalRows As Long
    Dim totalCols As Long
    'Dim ws As Worksheet
    Dim multiFilterCount As Integer
    
    Dim BetweenHeaders As String
    Dim BetweenHeadersArr() As String
    Dim inBetween() As String
    Dim inBetweenVals As String
    Dim inBetweenValsArr() As String
    Dim BetweenHeadersCount As String
    Dim FinalFiltersArr() As Variant
    
    On Error GoTo Handler:
    
    'Set ws = Sheet2
    totalRows = 1
    totalCols = 1
    multiFilterCount = 0
    BetweenHeadersCount = 0
    
    
    
    ReDim FinalFiltersArr(1 To totalRows, 1 To totalCols)
    
    For Each rng In ws.Range("D14:T14")
    
        If rng.Value <> "" Then
            
            filtPrint ws, rng.Offset(-1)
            
            'On Error GoTo 0
            If rng.Offset(-1).Value = "Between" Then
            
                If BetweenHeaders = "" Then
                
                    BetweenHeaders = rng.Offset(3).Value
                    
                    inBetween = Split(rng.Offset(2), ",")
                    
                    inBetweenVals = Replace(inBetween(0), "Txt", rng.Value) & "," & Replace(inBetween(1), "Txt", rng.Offset(1).Value)
                    
                    BetweenHeadersCount = BetweenHeadersCount + 1
                    
                Else
                    
                    BetweenHeaders = BetweenHeaders & "," & rng.Offset(3).Value
                    
                    ReDim inBetween(1 To 2)
                    
                    inBetween = Split(rng.Offset(2), ",")
                    
                    inBetweenVals = inBetweenVals & "," & Replace(inBetween(0), "Txt", rng.Value) & "," & Replace(inBetween(1), "Txt", rng.Offset(1).Value)
                    
                    BetweenHeadersCount = BetweenHeadersCount + 1
                
                End If
            
            ElseIf InStr(1, rng.Value, ";", vbTextCompare) > 0 Then
            
                ReDim filters(1 To 30)
                
                filters = Split(rng.Value, ";")
                
                For i = LBound(filters) To UBound(filters)
                
                    If filtHeaders = "" Then
                        
                        filtHeaders = rng.Offset(3).Value
                                            
                        totalCols = totalCols + 1
                    
                    Else
                    
                        filtHeaders = filtHeaders & "," & rng.Offset(3).Value
                                        
                        totalCols = totalCols + 1
                    
                    End If
                
                    If filtervals = "" Then
                    
                        filtervals = Replace(rng.Offset(2).Value, "Txt", filters(i))
                    
                    Else
                    
                        filtervals = filtervals & ";" & Replace(rng.Offset(2).Value, "Txt", filters(i))
                    
                    End If
                
                Next i
                
                
                multiFilterCount = multiFilterCount + 1
                
                totalRows = totalRows * (UBound(filters) + 1)
                
            Else
                    
                If filtHeaders = "" Then
                
                    filtHeaders = rng.Offset(3).Value
                    
                    filtervals = Replace(rng.Offset(2).Value, "Txt", rng.Value)
                    
                    totalCols = totalCols + 1
                
                Else
                
                    filtHeaders = filtHeaders & "," & rng.Offset(3).Value
                    
                    filtervals = filtervals & ";" & Replace(rng.Offset(2).Value, "Txt", rng.Value)
                    
                    totalCols = totalCols + 1
                    
                End If
                
            End If
        
        End If
    
    
    
    Next rng
    
    Dim filtersHeaderArr() As String, countHeaders As Integer, k As Long, a As Long, previousCount As Integer
    a = 3
    
    filterValsArr = Split(filtervals, ";")
    filtersHeaderArr = Split(filtHeaders, ",")
    
    previousCount = 1
    
    resultsheet.Cells(1, db_filters).CurrentRegion.ClearContents
    If UBound(filtersHeaderArr) >= 0 Then
    
        ReDim FinalFiltersArr(1 To totalRows, 1 To totalCols + BetweenHeadersCount * 2)
        
        resultsheet.Cells(1, db_filters).Resize(1, UBound(filtersHeaderArr) + 1).Value = filtersHeaderArr
    
        Dim l As Long, filtcol As Long
        
        For i = LBound(filtersHeaderArr) To UBound(filtersHeaderArr)
            
            countHeaders = 0
            
            For j = LBound(filtersHeaderArr) To UBound(filtersHeaderArr)
                
                If filtersHeaderArr(i) = filtersHeaderArr(j) Then countHeaders = countHeaders + 1
                
            Next j
            
            previousCount = previousCount * countHeaders
        
            
            a = 2
            
            If countHeaders > 1 Then
                
                l = 2
                
                Do While (l <= totalRows + 1)
                
                    For j = i To i + countHeaders - 1
                        
                        
                        For k = a To a + (totalRows / previousCount) - 1
                            
                            FinalFiltersArr(l - 1, j + 1) = "=" & Chr(34) & filterValsArr(j) & Chr(34)
                            
                        l = l + 1
                        
                        Next k
                       a = a + (totalRows / previousCount)
                      
                    Next j
    
                Loop
                
            Else
            
                For j = i To i + countHeaders - 1
                    
                    
                    For k = a To a + (totalRows / countHeaders) - 1
                    
                        FinalFiltersArr(k - 1, j + 1) = "=" & Chr(34) & filterValsArr(j) & Chr(34)
    
                    
                    Next k
                   a = a + (totalRows / countHeaders)
                    
                    
                Next j
            
            End If
                   
            
            i = i + countHeaders - 1
          
        Next i
    
    End If
    Dim betweenFilterCol As Long
    
    
    betweenFilterCol = resultsheet.Cells(1, db_filters).Offset(0, resultsheet.Cells(1, db_filters).CurrentRegion.Columns.count).Column
    
    If UBound(FinalFiltersArr, 2) <= 1 Then ReDim Preserve FinalFiltersArr(1 To UBound(FinalFiltersArr), 1 To totalCols + BetweenHeadersCount * 2)
    
    If BetweenHeadersCount <> 0 And BetweenHeadersCount = 1 Then
        
        
        inBetweenValsArr = Split(inBetweenVals, ",")
        resultsheet.Cells(1, betweenFilterCol).Value = BetweenHeaders
        resultsheet.Cells(1, betweenFilterCol + 1).Value = BetweenHeaders
        For i = LBound(FinalFiltersArr) To UBound(FinalFiltersArr)
        
            FinalFiltersArr(i, totalCols) = inBetweenValsArr(0)
            FinalFiltersArr(i, totalCols + 1) = inBetweenValsArr(1)
            
        Next i
        
    ElseIf BetweenHeadersCount <> 0 And BetweenHeadersCount > 1 Then
        
        BetweenHeadersArr = Split(BetweenHeaders, ",")
        
        For i = 0 To BetweenHeadersCount - 1
            
            inBetweenValsArr = Split(inBetweenVals, ",")
            resultsheet.Cells(1, betweenFilterCol).Value = BetweenHeadersArr(i)
            resultsheet.Cells(1, betweenFilterCol + 1).Value = BetweenHeadersArr(i)
               
            For j = LBound(FinalFiltersArr) To UBound(FinalFiltersArr)
            
                FinalFiltersArr(j, betweenFilterCol - db_filters + 1) = inBetweenValsArr(i * 2)
                FinalFiltersArr(j, betweenFilterCol - db_filters + 2) = inBetweenValsArr(i * 2 + 1)
            
            Next j
               
            betweenFilterCol = betweenFilterCol + 2
        
        Next i
    
    End If
    
    resultsheet.Cells(2, db_filters).Resize(UBound(FinalFiltersArr, 1), UBound(FinalFiltersArr, 2)).Value = FinalFiltersArr




Exit Sub

Handler:
resultsheet.Range("DA1").CurrentRegion.ClearContents

End Sub



Sub ApplyFilters(ws As Worksheet)
    'Procedure to Populate User entered filters and The pre-Defined filters to fill the respective pages Based on the Type of contract
    Dim i As Long
    Dim j As Integer, k As Integer, l As Integer, m As Integer
    Dim col As Integer
    Dim cnt As Integer
    Dim filtRow As Integer
    Dim target As Range
    Dim sh As Worksheet
    Dim filterTxt As String, filters() As String
    Dim multifilter As Integer
    Dim limitFilter As Integer
    Dim multiFiltHeader As String
    Dim multiFiltList() As String
    Dim mastercount As Double
    Dim Settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff

    unprotc ws
    
    ThisWorkbook.Activate
        'If user is not PCO filter and show results of all contracts ie choose the base data from database
    If ws.name <> Sheet16.name And Left(Sheet12.Range("Position"), 3) = "PCO" And ws.name <> Sheet19.name Then
        'If user is in PCO position choose contracts from myContracts table
        Set sh = Sheet14
    
    Else
        Set sh = Sheet8
    
    End If
    
    '
    
    filtRow = pg_FilterRow
    
      
    
    sh.Cells(1, db_filters).CurrentRegion.ClearContents
    
    mastercount = 0
    
    col = db_filters
    
    For i = 4 To 20
        
        Set target = ws.Cells(filtRow, i)
    
        If ws.name = Sheet2.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Contains"
            
            ws.Cells(filtRow + 1, i).Value = "Active"
        
        ElseIf ws.name = Sheet20.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Equals"
            
            ws.Cells(filtRow + 1, i).Value = "Closed"
        
        ElseIf ws.name = Sheet4.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Contains"
            
            ws.Cells(filtRow + 1, i).Value = "Amendment"
        
        ElseIf ws.name = Sheet3.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Ends With"
            
            ws.Cells(filtRow + 1, i).Value = "(No Existing)"
        
        ElseIf ws.name = Sheet5.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Ends With"
            
            ws.Cells(filtRow + 1, i).Value = "(Replace Existing)"
        
        ElseIf ws.name = Sheet10.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Equals"
            
            ws.Cells(filtRow + 1, i).Value = "Deviation"
        
        ElseIf ws.name = Sheet23.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Contains"
            
            ws.Cells(filtRow + 1, i).Value = "Renewal"
        
        ElseIf ws.name = Sheet24.name And ws.Cells(filtRow + 4, i).Value = "Type" Then
            
            target.Value = "Ends With"
            
            ws.Cells(filtRow + 1, i).Value = "Extension"
              
        End If
        
        filtPrint ws, target.Offset(1)
    
skipHere:
    Next i
    
    printFilters ws, sh


    Dim filtRng As Range
    Dim rng As Range, rng2 As Range, temprng As Range

    If ws.name = Sheet19.name Then
        
        Set filtRng = Sheet8.Range("DA1").CurrentRegion
    
        
        If Sheet8.Range("DA1").Value = "" Then
            
            Set rng = Sheet8.Range("DA1")
            
            rng.Value = "Temp PCO1"
            rng.Offset(1, 0).Value = Sheet12.Range("pName").Value
    
            rng.Offset(0, 1).Value = "Temp PCO2"
            rng.Offset(2, 1).Value = Sheet12.Range("pName").Value
            
        Else
        
            Set rng = Sheet8.Range("DA1").Offset(0, filtRng.Columns.count)
                
                rng.Value = "Temp PCO1"
                rng.Offset(0, 1).Value = "Temp PCO2"
                                                
                For Each rng2 In filtRng
                
                If rng2.row <> 1 Then
                    
                    If rng2.Value <> "" Then
                        
                        Sheet8.Range(rng2.Address).Offset(filtRng.rows.count - 1) = rng2
                        
                                     
                    End If
                   
                End If
                
                Next rng2
                
                
                For j = 2 To Application.WorksheetFunction.RoundUp(((Sheet8.Range("DA1").CurrentRegion.rows.count) / 2), 0)
            
                    rng.Offset(j - 1, 0).Value = Sheet12.Range("pName").Value
                    
                    rng.Offset(filtRng.rows.count - 2 + j, 1).Value = Sheet12.Range("pName").Value
                Next j
            
            
            End If
                
    End If
    
    On Error GoTo 0
    
    protc ws
    
    Settings.Restore

End Sub

Sub testArray()
    Dim TsdArr() As Variant
    Dim dbTable As ListObject
    
    Set dbTable = Sheet8.ListObjects(1)
    
    TsdArr = Application.Transpose(dbTable.ListColumns(db_startDate).DataBodyRange.Value)


End Sub

Sub AllDatesCalc()

    Dim ted               As Date 'term End Date
    Dim Med             As Date ' Max end date
    Dim Nor              As Integer ' No. Of Renewals
    Dim Nrd              As Date 'Next renewal dates
    Dim Erd              As Double ' each renewal duration
    Dim Tsd              As Date 'Term start date
    Dim Tdy              As Date
    Dim renPeriod As String
    Dim Dlfr As Variant
    Dim contrNum        As String
    Dim j               As Long, i As Long
    Dim k               As Long
    Dim Yesno           As String
    Dim TsdArr() As Variant
    Dim TedArr() As Variant
    Dim MedArr() As Variant
    Dim NorArr() As Variant
    Dim ErdArr() As Variant
    Dim NrdArr() As Variant
    Dim currRenPeriodArr() As String
    Dim dlfrArr() As Variant
    Dim linkedContracts() As Variant
    Dim contrNumArr() As Variant
    
    Dim sh As Worksheet
    
    Set sh = Sheet8
    
    Dim dbTable As ListObject
    
    Set dbTable = Sheet8.ListObjects(1)
    
    TsdArr = Application.Transpose(dbTable.ListColumns(db_startDate).DataBodyRange.Value)
    TedArr = Application.Transpose(dbTable.ListColumns(db_EndDate).DataBodyRange.Value)
    MedArr = Application.Transpose(dbTable.ListColumns(db_MaxendDate).DataBodyRange.Value)
    NorArr = Application.Transpose(dbTable.ListColumns(db_NoOfRens).DataBodyRange.Value)
    ErdArr = Application.Transpose(dbTable.ListColumns(db_EachRenDur).DataBodyRange.Value)
    linkedContracts = Application.Transpose(dbTable.ListColumns(db_ContrLinked).DataBodyRange.Value)
    contrNumArr = Application.Transpose(dbTable.ListColumns(db_contrNumber).DataBodyRange.Value)

    
    For i = LBound(TsdArr) To UBound(TsdArr)
    
    Tsd = TsdArr(i)
    
    ted = TedArr(i)
    
    Med = MedArr(i)
    
    Nor = NorArr(i)
    
    Erd = ErdArr(i)
    
    Tdy = Now
    
    If Med = 0 And ted <> 0 Then

        Med = ted + Nor * Erd * 365
            
    End If

    
    Nrd = NextRenDate(Erd, ted, Nor, Med, Tsd)
    renPeriod = CurrRenPeriod(Tdy, Tsd, Med, ted, Nor, Erd)
    
    If renPeriod = "Closed" Or renPeriod = vbNullString Or Nrd = 0 Then
        
        Dlfr = vbNullString
        
    ElseIf Tdy > ted Then
    
        Dlfr = DateDiff("d", Tdy, Nrd)
      
    Else
    
        Dlfr = DateDiff("d", Tdy, ted)
    
    End If
    
    If i = 1 Then
        
        ReDim NrdArr(1 To 1)
        NrdArr(1) = Nrd
        
        ReDim currRenPeriodArr(1 To 1)
        currRenPeriodArr(1) = renPeriod
        
        ReDim dlfrArr(1 To 1)
        dlfrArr(1) = Dlfr
        
    Else
        
        ReDim Preserve NrdArr(1 To UBound(NrdArr) + 1)
        NrdArr(UBound(NrdArr)) = IIf(Nrd < DateSerial(2000, 1, 1), vbNullString, Nrd)
        
        ReDim Preserve currRenPeriodArr(1 To UBound(currRenPeriodArr) + 1)
        currRenPeriodArr(UBound(currRenPeriodArr)) = renPeriod
        
        ReDim Preserve dlfrArr(1 To UBound(dlfrArr) + 1)
        dlfrArr(UBound(dlfrArr)) = Dlfr
        
    End If
    
    Next i

    dbTable.ListColumns(db_NextRenDate).DataBodyRange.Value = Application.Transpose(NrdArr)
    dbTable.ListColumns(db_CurrRenPeriod).DataBodyRange.Value = Application.Transpose(currRenPeriodArr)
    dbTable.ListColumns(db_DaysLeftFrRen).DataBodyRange.Value = Application.Transpose(dlfrArr)
    
    For i = LBound(linkedContracts) To UBound(linkedContracts)
    
        If linkedContracts(i) <> "" Then
            
            For j = LBound(contrNumArr) To UBound(contrNumArr)
                
                If contrNumArr(j) = linkedContracts(i) Then
                    
                    sh.Cells(i + 1, db_DaysLeftFrRen).Value = sh.Cells(j + 1, db_DaysLeftFrRen).Value
                    
                    Exit For
                
                End If
            
            Next j
        
        End If
        
    Next i
    


End Sub

Sub CalcDates(sh As Worksheet, i As Long)

    Dim ted               As Date 'term End Date
    Dim Med             As Date ' Max end date
    Dim Nor              As Integer ' No. Of Renewals
    Dim Nrd              As Date 'Next renewal dates
    Dim Erd              As Double ' each renewal duration
    Dim Tsd              As Date 'Term start date
    Dim Tdy              As Date
    Dim contrNum        As String
    Dim j               As Long
    Dim k               As Long
    Dim Yesno           As String
    Dim TsdArr() As Variant
    Dim TedArr() As Variant
    Dim MedArr() As Variant
    Dim NorArr() As Integer
    Dim ErdArr() As Double
    Dim NrdArr() As Variant
    Dim dbTable As ListObject

    
'    For i = LBound(TsdArr) To UBound(Tsd)
'
'    Next i
    
    Tsd = sh.Cells(i, db_startDate).Value
    
    ted = sh.Cells(i, db_EndDate).Value
    
    Med = sh.Cells(i, db_MaxendDate).Value
    
    Nor = sh.Cells(i, db_NoOfRens).Value
    
    Erd = sh.Cells(i, db_EachRenDur).Value
    
    Tdy = Now
    
    Nrd = NextRenDate(Erd, ted, Nor, Med, Tsd)
    
    sh.Cells(i, db_NextRenDate).Value = Nrd
    
    If Med = 0 And ted <> 0 Then

        Med = ted + Nor * Erd * 365
        
        sh.Cells(i, db_MaxendDate).Value = Med
    
    End If
    
    If Tdy > ted Then
        
        Nrd = NextRenDate(Erd, ted, Nor, Med, Tsd)
                            
        sh.Cells(i, db_DaysLeftFrRen).Value = DateDiff("d", Tdy, Nrd)
    
        'sh.Cells(i,  db_DaysLeftFrRen).Value = 0
            
    Else
        
        sh.Cells(i, db_DaysLeftFrRen).Value = DateDiff("d", Tdy, ted)
    
    End If
    
    sh.Cells(i, db_CurrRenPeriod).Value = CurrRenPeriod(Tdy, Tsd, Med, ted, Nor, Erd)
    
    If sh.Cells(i, db_CurrRenPeriod).Value = vbNullString Then sh.Cells(i, db_DaysLeftFrRen).Value = vbNullString
    
    If sh.Range("C" & i).Value = "Closed" Then
                
        sh.Cells(i, db_Priority).Value = "Low"
        
        sh.Cells(i, db_DaysLeftFrRen).Value = vbNullString
               
    End If
                
    'Debug.Print sh.Cells(i,  db_DaysLeftFrRen).Value
    
    If sh.Cells(i, db_DaysLeftFrRen).Value < 100 And sh.Cells(i, db_DaysLeftFrRen).Value <> "0" Then
        
        If sh.Cells(i, db_DaysLeftFrRen).Value <> vbNullString And sh.Cells(i, db_Priority).Value <> "High" Then
        
            Yesno = MsgBox("Days Left For Renewal is less than 100 days!" & vbNewLine & "Mark as High Priority?", vbYesNo, "Mark as High priority")

            If Yesno = vbYes Then sh.Cells(i, db_Priority).Value = "High"
        
        End If
    
    ElseIf sh.Cells(i, db_DaysLeftFrRen).Value < 150 And sh.Cells(i, db_Priority).Value = "" And sh.Cells(i, db_DaysLeftFrRen).Value <> 0 Then
        
        sh.Cells(i, db_Priority).Value = "Medium"
    
    ElseIf sh.Cells(i, db_Priority).Value = "" And sh.Cells(i, db_DaysLeftFrRen).Value <> 0 Then
        
        sh.Cells(i, db_Priority).Value = "Low"
    
    End If


    If sh.Cells(i, 32).Value <> "" Then
        
        contrNum = sh.Cells(i, 32).Value
                    
        For j = 2 To Sheet8.Cells(1, 1).End(xlDown).row
            
            If sh.Cells(j, 4).Value = contrNum Then
                
                sh.Cells(i, db_DaysLeftFrRen).Value = sh.Cells(j, db_DaysLeftFrRen).Value
                
                Exit For
            
            End If
        
        Next j
    
    End If


End Sub




Sub applyAdvFilt(Optional newSheet As Worksheet)
    
    'Dim Timer As New TimerCls
    Dim MasterTimer As New TimerCls
    Dim Settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff
    MasterTimer.start
    'Timer.start
    
    Dim ted             As Date 'term End Date
    Dim Med             As Date ' Max end date
    Dim Nor             As Integer ' No. Of Renewals
    Dim Nrd             As Date 'Next renewal dates
    Dim Erd             As Double ' each renewal duration
    Dim Tsd             As Date 'Term start date
    Dim Tdy             As Date
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
    
    If newSheet Is Nothing Then
        
        Set parentsheet = ThisWorkbook.ActiveSheet
    
    Else
        
        Set parentsheet = newSheet
    
    End If
    
    
    unprotc parentsheet
    
    parentsheet.Range("A9").Value = True
    
    If Left(Sheet12.Range("Position"), 3) = "PCO" And parentsheet.name <> Sheet16.name And parentsheet.name <> Sheet19.name Then
        
        Set sh = Sheet14
        
        Sheet14.Range("DA1").Value = "PCO"
        
        Sheet14.Range("DA2").Value = Range("pName").Value
        
        Set rng = Sheet8.Range("A1").CurrentRegion
        
        Set src = Sheet14.Range("DA1").CurrentRegion
        
        Set drng = Sheet14.Range("A1").CurrentRegion

        sh.ListObjects(1).Resize rng
        
        rng.AdvancedFilter xlFilterCopy, src, drng
        
        If drng.CurrentRegion.rows.count <> 1 Then sh.ListObjects(1).Resize drng.CurrentRegion
    
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
    
    'Timer.PrintTime "Advanced Filter"
    
    resultcolCount = 1
    
    rows = sh.Range("BA1").End(xlDown).row - 1
    
    parentsheet.Range(parentsheet.ListObjects(1)).ClearContents
    
    clrCondFormat Range(parentsheet.ListObjects(1))
    
    sourceArr = Application.Transpose(sh.Range("BA1").CurrentRegion.Value)
    
    col = parentsheet.Range("D17").End(xlToRight).Column
    
    ReDim DestinArr(1 To col, 1 To UBound(sourceArr, 2))

    For i = 4 To col

        For j = LBound(sourceArr, 1) To UBound(sourceArr, 1)

            If parentsheet.Cells(17, i).Value = sourceArr(j, 1) Then

                For k = LBound(sourceArr, 2) + 1 To UBound(sourceArr, 2)

                    DestinArr(resultcolCount, k - 1) = sourceArr(j, k)
                
                Next k

                
                resultcolCount = resultcolCount + 1
            End If

        Next j

        If parentsheet.Cells(17, i).Value = "Days Left For Renewal/Expiry" Then dlfrFinder = i

    Next i
    'Sheet26.Cells(1, 1).Resize(UBound(DestinArr, 2), UBound(DestinArr, 1)).Value = Application.Transpose(DestinArr)
    parentsheet.Cells(18, 4).Resize(UBound(DestinArr, 2), UBound(DestinArr, 1)).Value = Application.Transpose(DestinArr)
    
    'Timer.PrintTime "Filter Columns"

skipHere:
    
    Dim addr As String
    
    addr = parentsheet.Cells(17, Columns.count).End(xlToLeft).Address
    
    If parentsheet.Range("D17").End(xlToRight).Column >= 26 Then
        
        addr = Left(addr, 4)
    
    Else
        
        addr = Left(addr, 3)
    
    End If
    
    'On Error GoTo Handler
    
    unprotc parentsheet

On Error Resume Next
    parentsheet.ListObjects(1).DataBodyRange.RowHeight = 25
    
    
    
    If Err.Number <> 0 Or parentsheet.ListObjects(1).ListRows.count <= 7 Then parentsheet.Range("D18:D24").RowHeight = 25
On Error GoTo 0
    parentsheet.Range(parentsheet.ListObjects(1).name & "[Description]").ColumnWidth = 50
    
    With parentsheet.ListObjects(1).Range
    
        .RowHeight = 25
    
        .VerticalAlignment = xlVAlignCenter
    
        .HorizontalAlignment = xlHAlignCenter

    End With

    parentsheet.Range("16:16").EntireRow.Hidden = True

    On Error Resume Next
    
    parentsheet.ListObjects(1).Resize parentsheet.Range("D17:" & addr & rows + 17)
    
    If Err.Number <> 0 And parentsheet.Range("E18") = "" Then parentsheet.ListObjects(1).Resize parentsheet.Range("D17:" & addr & "18")
    
    On Error GoTo 0

    parentsheet.Range(parentsheet.ListObjects(1).name).Locked = False

    dlfrDatabar parentsheet, parentsheet.ListObjects(1).ListRows.count + 17, dlfrFinder

    rowHighlight parentsheet.Range(parentsheet.ListObjects(1).name)

    priorityConditions parentsheet

    Range(parentsheet.ListObjects(1).name & "[#Headers]").Locked = False

    
    If Not parentsheet.ListObjects(1).ShowAutoFilter Then Range(parentsheet.ListObjects(1).name & "[#Headers]").AutoFilter

    parentsheet.Range("A9").Value = False

    protc parentsheet
    
    'Timer.PrintTime "Apply Formatting"
    
    MasterTimer.PrintTime "ApplyAdvFilter"
    
    Set sh = Nothing
    
    Set parentsheet = Nothing
    
    Application.screenUpdating = True
    
    Dim errcount As Integer
    
    Settings.Restore
    
    Exit Sub
    
Handler:
    
    errcount = errcount + 1
        
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff

    
    If errcount = 1 Then
    
        updateLog ThisWorkbook, Err.Number & ":" & Err.Description & " " & parentsheet.name, "ApplyadvFilter Failed, trying again once more"
        
        '
        
        'parentsheet.Select
        
        applyAdvFilt parentsheet
        
    Else
    
        updateLog ThisWorkbook, Err.Number & ":" & Err.Description & " " & parentsheet.name, "ApplyadvFilter Failed, tryied once and failed again"
        
        parentsheet.Select
    
    End If
    
    Settings.Restore

End Sub


Sub testFormatting()

unprotc ActiveSheet
    dlfrDatabar ActiveSheet, 169, 10
    rowHighlight ActiveSheet.Range(ActiveSheet.ListObjects(1).name)
    priorityConditions ActiveSheet

protc ActiveSheet
End Sub

Function IsListobjectFiltered(ByVal listObj As ListObject) As Boolean

    If listObj.ShowAutoFilter Then
        'If listObj.AutoFilter.FilterMode Then
            IsListobjectFiltered = True
            Exit Function
        'End If
    End If

    IsListobjectFiltered = False

End Function


Sub priorityConditions(ws As Worksheet)
'Few Conditions to Auto pupulate Priority settings based on Days left for renewal
    If ws Is Nothing Then
        
        Set ws = ActiveSheet
    
    End If
    
    Dim rng As Range
    
    Set rng = ws.Range(ws.ListObjects(1).name & "[Priority]")
'rng.FormatConditions.Delete
'Low
    unprotc ws
    
    rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D18=""Low"""
    
    rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        
        .Color = -11489280
        
        .TintAndShade = 0
    
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    Application.CutCopyMode = False
'Medium
    
    rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D18=""Medium"""
    
    rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        
        .ThemeColor = xlThemeColorAccent2
        
        .TintAndShade = 0
    
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    Application.CutCopyMode = False
'High
    
    rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D18=""High"""
    
    rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        
        .Color = -16776961
        
        .TintAndShade = 0
    
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    Application.CutCopyMode = False
    
    If ws.name <> Sheet3.name Then Set rng = ws.Range(ws.ListObjects(1).name & "[Days Left For Renewal/Expiry]")

'rng.FormatConditions.Delete
'Low
    rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D18=""Low"""
    
    rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        
        .Color = -11489280
        
        .TintAndShade = 0
    
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    Application.CutCopyMode = False
'Medium
    
    rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D18=""Medium"""
    
    rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        
        .ThemeColor = xlThemeColorAccent2
        
        .TintAndShade = 0
    
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    Application.CutCopyMode = False
'High
    
    rng.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$D18=""High"""
    
    rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
    
    With rng.FormatConditions(1).Font
        
        .Color = -16776961
        
        .TintAndShade = 0
    
    End With
    
    rng.FormatConditions(1).StopIfTrue = False
    
    Application.CutCopyMode = False

End Sub



Sub reframe()

Dim sh As Worksheet
Dim Settings As New ExclClsSettings

Settings.TurnOff

For Each sh In ThisWorkbook.Worksheets
    If sh.Range("A1").Value = "NavTo" Then
        unprotc sh
            '
            Navto sh
            sh.Range("A7:C8").Locked = False
        protc sh
    End If
Next sh
Settings.Restore

End Sub




Sub rowHighlight(rng As Range)

    With rng
        
        On Error Resume Next
        
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$A$7=ROW()"
        
        .FormatConditions(rng.FormatConditions.count).SetFirstPriority
    
    End With
    
    With rng.FormatConditions(1).Font
        
        .Bold = True
        
        .Italic = False
        
        .ThemeColor = xlThemeColorAccent1
        
        .TintAndShade = -0.499984740745262
    
    End With
    
    With rng.FormatConditions(1).Interior
        
        .PatternColorIndex = xlAutomatic
        
        '.Color = Sheet11.Range("InactiveBtn_Mid").Interior.Color
        
        .ThemeColor = xlThemeColorDark1
        
        '.TintAndShade = -0.249946592608417
    
    End With
    
    rng.FormatConditions(1).StopIfTrue = False

End Sub



Sub clrCondFormat(rng As Range)
''
    With rng
        
        unprotc ThisWorkbook.Worksheets(rng.Parent.name)
        
        .FormatConditions.Delete
        
        .Interior.ColorIndex = 0
    End With

End Sub




Sub dlfrDatabar(Optional sheet As Worksheet, Optional newLastrow As Long = 1, Optional dlfrFinder As Integer)

    Dim Settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff

    Dim rng As Range
    
    If newLastrow <> 0 Then
    
    On Error Resume Next
    
    Set rng = sheet.Range(sheet.Cells(18, dlfrFinder), sheet.Cells(newLastrow, dlfrFinder))
    
    sheet.Range("I1:I17").FormatConditions.Delete
    
    If sheet.name = Sheet16.name Then
        
        sheet.Range("J1:J17").FormatConditions.Delete
    
    End If
    
    rng.FormatConditions.Delete
        
        ''
        
        ''
        
        rng.FormatConditions.AddDatabar
        
        rng.FormatConditions(rng.FormatConditions.count).ShowValue = True
        
        rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
        
        With rng.FormatConditions(1)
            
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        
        End With
        
        With rng.FormatConditions(1).BarColor
            
            .ThemeColor = xlThemeColorAccent6
            
            .TintAndShade = 0.399975585192419
        
        End With
        
        rng.FormatConditions(1).BarFillType = xlDataBarFillGradient
        
        rng.FormatConditions(1).Direction = xlContext
        
        rng.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        
        rng.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
        
        rng.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        
        With rng.FormatConditions(1).AxisColor
            
            .Color = 0
            
            .TintAndShade = 0
        
        End With
        
        With rng.FormatConditions(1).NegativeBarFormat.Color
            
            .Color = 255
            
            .TintAndShade = 0
        
        End With
        
        rng.FormatConditions.AddColorScale ColorScaleType:=2
        
        rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
        
        rng.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueLowestValue
        
        With rng.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            
            .ThemeColor = xlThemeColorLight1
            
            .TintAndShade = 0.349986266670736
        
        End With
        
        rng.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueHighestValue
        
        With rng.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            
            .ThemeColor = xlThemeColorDark2
            
            .TintAndShade = -0.899990844447157
        
        End With
            
            With rng.Offset(1).Font
            
            .ThemeColor = xlThemeColorAccent5
            
            .TintAndShade = -0.249977111117893
        
        End With
        
        rng.Font.Bold = True
        
        rng.Font.Underline = xlUnderlineStyleNone
        
        rng.Font.Size = 12
        
        With rng
            
            .HorizontalAlignment = xlRight
            
            .VerticalAlignment = xlCenter
        
        End With
        
        rng.NumberFormat = "General"
    
    End If
    
    Settings.Restore
 
End Sub



Function NextRenDate(Erd As Double, ted As Date, Nor As Integer, Med As Date, Tsd As Date) As Date

Dim yr As Long, m As Integer, d As Integer, i
Dim currPeriod As String

If Med = 0 And ted <> 0 Then

    Med = ted + Nor * Erd * 365

End If

currPeriod = CurrRenPeriod(Now, Tsd, Med, ted, Nor, Erd)
    
    If currPeriod = "Initial Term" Or currPeriod = "Yet to Start" Then
        
        NextRenDate = ted
        
        Exit Function
    
    ElseIf currPeriod = "Closed" Or currPeriod = vbNullString Then
        
        If Med <> 0 Then
            NextRenDate = Med
        Else
            NextRenDate = 0
        End If
        Exit Function
    
    ElseIf currPeriod = "Extension with no Renewals" Or currPeriod = "Extension" Then
        
        NextRenDate = Med
        
        Exit Function
    
    End If
    
    If Erd <> 0 Then yr = Year(ted) + Application.WorksheetFunction.RoundUp((DateDiff("d", ted, Now) / 365), 0)
    
    If Erd = 0 Then
        
        yr = Year(Med)
    
    End If
    
    m = Month(ted)
    
    d = Day(ted)
    
    'yr =
    '
    
    If DateSerial(yr, m, d) >= Med Then
        
        NextRenDate = Med
    
    Else
        
        For i = 1 To Nor
        
            If DateSerial(Year(ted) + i * Erd, Month(ted), Day(ted)) > Now Then
            
                NextRenDate = DateSerial(Year(ted) + i * Erd, Month(ted), Day(ted))
                
                Exit Function
                
                Exit For
            
            End If
        
        Next i
        
    
    End If

End Function




Function CurrRenPeriod(Tdy As Date, Tsd As Date, Med As Date, ted As Date, Nor As Integer, Erd As Double) As String

Dim contractEndDate As Integer

If Med = 0 And ted <> 0 Then

    Med = ted + Nor * Erd * 365

End If

If ted = 0 Or Med = 0 Then
    
    CurrRenPeriod = vbNullString

ElseIf Tsd > Tdy Then
    
    CurrRenPeriod = "Yet to Start"

ElseIf Tdy > Med Then
    
    CurrRenPeriod = "Closed"

ElseIf Tdy < ted Then
    
    CurrRenPeriod = "Initial Term"

ElseIf Tdy > ted And Nor = 0 And ted < Med Then
    
    CurrRenPeriod = "Extension with no Renewals"

ElseIf Tdy > ted And Nor > 0 And Application.WorksheetFunction.RoundUp(DateDiff("d", ted, Tdy) / (365 * Erd), 0) > Nor And ted <> Med Then
    
    CurrRenPeriod = "Extension"

Else
    
    CurrRenPeriod = Application.WorksheetFunction.IfError("RENEWAL " & Application.WorksheetFunction.RoundUp(DateDiff("d", Tdy, ted) / (365 * Erd), 0), "")

End If

End Function

Sub refreshAllContracts()
    
    Dim Settings As New ExclClsSettings
    Dim stats As New statClass
    
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    
    Settings.TurnOff

    unProtectWorksheet
    
    Dim i As Long
    
    Dim col As Long
    
    Dim rows As Long
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim j As Long
    Dim lastcol As Long
    
    Set sh = Sheet8
    
    Set ws = Sheet16
    
    stats.showStatus "Downloading Contracts Data.."
    
    getUpdatedData
    
    stats.showStatus "Downloading User Data.."
    
    SyncstaffData_FromGsheets "Mandatory"
    
    stats.showStatus "Downloading Field Access Data.."
    
    getFieldAccessData "Mandatory"
    
    col = ws.Range("D17").End(xlToRight).Column
    
    rows = sh.Range("A1").End(xlDown).row - 1
    
    lastcol = sh.Range("A1").End(xlToRight).Column
    
    ws.Range(ws.ListObjects(1).name).ClearContents
    
    stats.showStatus "Calculating Dates.."
    
    AllDatesCalc
    
    ConverttoDate
    
    sh.Range("FG1").Value = Now
        
    For i = 4 To col
        
        For j = 1 To lastcol
            
            If ws.Cells(17, i).Value = sh.Cells(1, j).Value Then
                
                ws.Range(ws.Cells(18, i).Address).Resize(rows, 1).Value = sh.Range(sh.Cells(2, j).Address).Resize(rows, 1).Value
            
            End If
        
        Next j
    
    Next i
    
    
    Dim addr As String
    
    addr = ws.Cells(17, ws.Columns.count).End(xlToLeft).Address
    
    If ws.Range("D17").End(xlToRight).Column >= 26 Then
        
        addr = Left(addr, 4)
    
    Else
        
        addr = Left(addr, 3)
    
    End If
    
    unprotc Sheet16
    
    ws.ListObjects(1).Resize ws.Range("D17:" & addr & rows + 17)
    
    Dim rng As Range, src As Range, drng As Range
    
    If Left(Sheet12.Range("Position"), 3) = "PCO" Then
    
        Sheet14.Range("DA1").Value = "PCO"
        Sheet14.Range("DA2").Value = Sheet12.Range("pName").Value
        Sheet14.Range("A1").Resize(1, Sheet8.ListObjects(1).ListColumns.count).Value = Sheet8.Range("A1").CurrentRegion.Resize(1, Sheet8.Range("A1").CurrentRegion.Columns.count).Value
    
        Set rng = Sheet8.Range("A1").CurrentRegion
        
        Set src = Sheet14.Range("DA1").CurrentRegion
        
        Set drng = Sheet14.Range("A1").CurrentRegion
        
        On Error Resume Next
        
            rng.AdvancedFilter xlFilterCopy, src, drng
        
        On Error GoTo 0
    
    End If
    
    stats.showStatus "Refreshing Data..."
    
    clearFilters ws
    
    protectWorksheet
    
    stats.closeStats
    
    Settings.Restore
    
End Sub

Sub updatePivotTable()

Dim srcTable As ListObject, destinTable As ListObject

Set srcTable = Sheet8.ListObjects(1)

Set destinTable = Sheet28.ListObjects(1)

destinTable.DataBodyRange.ClearContents

destinTable.DataBodyRange.Resize(srcTable.ListRows.count, srcTable.ListColumns.count).Value = srcTable.DataBodyRange.Value

destinTable.Resize Sheet28.Range("A1").CurrentRegion

End Sub

Sub RefDashboard()

    updatePivotTable
    
    Dim pt As PivotTable
    
    For Each pt In Sheet25.PivotTables

        pt.PivotCache.Refresh
    
    Next pt
    
    MsgBox "Refreshed!"

End Sub


Sub FindCode()

    Application.CommandBars("Edit").Controls("Find...").Execute

End Sub


Sub clearFilters(Optional ws As Worksheet)
    
    Dim rng As Range, cellRng As Range
    
    Dim notEmpty As Boolean
    
    notEmpty = False
    
    If ws Is Nothing Then
    
        Set ws = ActiveSheet
    
    End If
    
    Set rng = ws.Range("D13:T16")
    
        For Each cellRng In rng
        
        
        
        If cellRng.Value <> "" Then
        
            notEmpty = True
            
            Exit For
            
        End If
        
    Next cellRng
    
    rng.ClearContents
    
    On Error Resume Next
    
    ws.ListObjects(1).AutoFilter.ShowAllData
    
    ws.Range("I1").ClearContents
    

    If notEmpty = True Then applyAdvFilt ws
    
End Sub
