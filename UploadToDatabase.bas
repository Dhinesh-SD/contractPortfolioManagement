Attribute VB_Name = "UploadToDatabase"
Dim arrRowsTobeDeleted()
Dim intArrsize
Dim deletedFlag As Boolean
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

'Admin
'https://docs.google.com/forms/d/e/1FAIpQLSceZtykKSeGb5KRRHlYneP7-eynkk81UyeCQJ_KWMP_CCY2qg/viewform?usp=pp_url
'PCO_1
'https://docs.google.com/forms/d/e/1FAIpQLSf-l4djrWGInh_7OSqV2RbsaSNegkHWfgDKyDVeC-h5qzfxxg/viewform?usp=pp_url
'PCO_2
'https://docs.google.com/forms/d/e/1FAIpQLSeWLOiaXCyzMUYMwSiYy5c4bDg9Yw3zLtkXejK1eMX8xCDoiQ/viewform?usp=pp_url
'PCO_3
'https://docs.google.com/forms/d/e/1FAIpQLSeDj68Z76D_aNyi-SgWm1SntAt-OoMg0Z0EHNDNKRl8HG9nCA/viewform?usp=pp_url
'PCO_4
'https://docs.google.com/forms/d/e/1FAIpQLSchv13R2idfDyl_4PA0G76q_G4JLBn3rzghn3ttp3LjrvGXXg/viewform?usp=pp_url
'PCO_5
'https://docs.google.com/forms/d/e/1FAIpQLSceZtykKSeGb5KRRHlYneP7-eynkk81UyeCQJ_KWMP_CCY2qg/viewform?usp=pp_url
Sub syncData()

    Dim http As Object, http2 As Object
    Dim URL As String
    Dim backupURL As String, finalBackupURL As String
    Dim Primary_Key As String
    Dim pk_id As Long
    Dim PCO As String
    Dim typ As String, ContractNumber As String, CLMSNumber As String, doctype As String, RFx As String, Description As String
    Dim Division As String, DivisionContact As String, TempPco1 As String, TempPco2 As String, Status As String
    Dim Amd_1 As String, Amd_2 As String, Amd_3 As String
    Dim DeviationType As String, DeviationReason As String, Agency As String, AgencyContact As String, Supplier As String
    Dim DeviationDate As String, StartDate As String, EndDate As String, Nor As String, Erd As String, MaxEndDate As String
    Dim Remarks As String, Notes As String, ExtensionDuration As String, EstimatedSpend As String, ContractLinked As String, Files As String
    Dim Priority As String, NextRenDate As String, DaysLforRen As String, CurrRenPeriod As String, CLMSReqNumber As String
    Dim SupplierABNumber As String, SharepointNumber As String, RecDateEntered As String, ContractDateEntered As String, RenewContract As String
    Dim Unique_ID As String, DeleteContract As String
    Dim i As Long
    Dim strData As String
    Dim finalURL As String
    Dim syncStatus As String
    Dim Security As String
    Dim position As String
    Dim ws As Worksheet
    
    Security = Sheet12.Range("Security")
    position = Sheet12.Range("Position")
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    Set http2 = CreateObject("MSXML2.ServerXMLHTTP")
    
    Set ws = Sheet8
    
    backupURL = "https://docs.google.com/forms/d/e/1FAIpQLSfGAPhYzuIyRySPWe7u__MUiuZvQhdv87iMvnDp54u9NCPcqw/formResponse?"

    If Security = "Admin" Then
        
        URL = "https://docs.google.com/forms/d/e/1FAIpQLSceZtykKSeGb5KRRHlYneP7-eynkk81UyeCQJ_KWMP_CCY2qg/formResponse?"
        
    ElseIf position = "PCO-1" Then
        
        URL = "https://docs.google.com/forms/d/e/1FAIpQLSf-l4djrWGInh_7OSqV2RbsaSNegkHWfgDKyDVeC-h5qzfxxg/formResponse?"
        
    ElseIf position = "PCO-2" Then
    
        URL = "https://docs.google.com/forms/d/e/1FAIpQLSeWLOiaXCyzMUYMwSiYy5c4bDg9Yw3zLtkXejK1eMX8xCDoiQ/formResponse?"
        
    ElseIf position = "PCO-3" Then
    
        URL = "https://docs.google.com/forms/d/e/1FAIpQLSeDj68Z76D_aNyi-SgWm1SntAt-OoMg0Z0EHNDNKRl8HG9nCA/formResponse?"

    ElseIf position = "PCO-4" Then
    
        URL = "https://docs.google.com/forms/d/e/1FAIpQLSchv13R2idfDyl_4PA0G76q_G4JLBn3rzghn3ttp3LjrvGXXg/formResponse?"

    ElseIf position = "PCO-5" Then
    
        URL = "https://docs.google.com/forms/d/e/1FAIpQLSceZtykKSeGb5KRRHlYneP7-eynkk81UyeCQJ_KWMP_CCY2qg/formResponse?"
    
    End If
    
    deletedFlag = False
    intArrsize = 0
    pk_id = Application.WorksheetFunction.Max(Sheet8.Range(Sheet8.ListObjects(1).name & "[Unique_ID]")) + 1
    
    For i = 2 To ws.Range("A1").End(xlDown).row
    
    syncStatus = ws.Range("AT" & i).Value
    
    If syncStatus = "" Then
    
    Primary_Key = ws.Range("A" & i).Value
    PCO = ws.Range("B" & i).Value
    typ = ws.Range("C" & i).Value
    ContractNumber = ws.Range("D" & i).Value
    CLMSNumber = ws.Range("E" & i).Value
    doctype = ws.Range("F" & i).Value
    RFx = ws.Range("G" & i).Value
    Description = ws.Range("H" & i).Value
    Division = ws.Range("I" & i).Value
    DivisionContact = ws.Range("J" & i).Value
    TempPco1 = ws.Range("K" & i).Value
    TempPco2 = ws.Range("L" & i).Value
    Status = ws.Range("M" & i).Value
    Amd_1 = ws.Range("N" & i).Value
    Amd_2 = ws.Range("O" & i).Value
    Amd_3 = ws.Range("P" & i).Value
    DeviationType = ws.Range("Q" & i).Value
    DeviationReason = ws.Range("R" & i).Value
    Agency = ws.Range("S" & i).Value
    AgencyContact = ws.Range("T" & i).Value
    Supplier = ws.Range("U" & i).Value
    DeviationDate = ws.Range("V" & i).Value
    StartDate = ws.Range("W" & i).Value
    EndDate = ws.Range("X" & i).Value
    Nor = ws.Range("Y" & i).Value
    Erd = ws.Range("Z" & i).Value
    MaxEndDate = ws.Range("AA" & i).Value
    Remarks = ws.Range("AB" & i).Value
    Notes = ws.Range("AC" & i).Value
    ExtensionDuration = ws.Range("AD" & i).Value
    EstimatedSpend = ws.Range("AE" & i).Value
    ContractLinked = ws.Range("AF" & i).Value
    Files = Replace(ws.Range("AG" & i).Value, "\", "%5C")
    Priority = ws.Range("AH" & i).Value
    NextRenDate = ws.Range("AI" & i).Value
    DaysLforRen = ws.Range("AJ" & i).Value
    CurrRenPeriod = ws.Range("AK" & i).Value
    CLMSReqNumber = ws.Range("AL" & i).Value
    SupplierABNumber = ws.Range("AM" & i).Value
    SharepointNumber = ws.Range("AN" & i).Value
    RecDateEntered = ws.Range("AO" & i).Value
    ContractDateEntered = ws.Range("AP" & i).Value
    RenewContract = ws.Range("AQ" & i).Value
    Unique_ID = ws.Range("AR" & i).Value
    DeleteContract = ws.Range("AS" & i).Value
    

    
    strData = "&entry.2069807768=" & Replace(Primary_Key, "#", "No.")
    strData = strData & "&entry.1039171754=" & Replace(PCO, "#", "No.")
    strData = strData & "&entry.1613953102=" & Replace(typ, "#", "No.")
    strData = strData & "&entry.2121600145=" & Replace(ContractNumber, "#", "No.")
    strData = strData & "&entry.1100413045=" & Replace(CLMSNumber, "#", "No.")
    strData = strData & "&entry.845829802=" & Replace(doctype, "#", "No.")
    strData = strData & "&entry.894141892=" & Replace(RFx, "#", "No.")
    strData = strData & "&entry.1204971453=" & Replace(Description, "#", "No.")
    strData = strData & "&entry.1773446982=" & Replace(Division, "#", "No.")
    strData = strData & "&entry.616441955=" & Replace(DivisionContact, "#", "No.")
    strData = strData & "&entry.1919545029=" & Replace(TempPco1, "#", "No.")
    strData = strData & "&entry.1912202854=" & Replace(TempPco2, "#", "No.")
    strData = strData & "&entry.1294908020=" & Replace(Status, "#", "No.")
    strData = strData & "&entry.1961247922=" & Replace(Amd_1, "#", "No.")
    strData = strData & "&entry.2031838576=" & Replace(Amd_2, "#", "No.")
    strData = strData & "&entry.923996207=" & Replace(Amd_3, "#", "No.")
    strData = strData & "&entry.263457955=" & Replace(DeviationType, "#", "No.")
    strData = strData & "&entry.801271685=" & Replace(DeviationReason, "#", "No.")
    strData = strData & "&entry.1328156162=" & Replace(Agency, "#", "No.")
    strData = strData & "&entry.1629641840=" & Replace(AgencyContact, "#", "No.")
    strData = strData & "&entry.347047104=" & Replace(Supplier, "#", "No.")
    strData = strData & "&entry.1392388851=" & Replace(DeviationDate, "#", "No.")
    strData = strData & "&entry.1474710651=" & Replace(StartDate, "#", "No.")
    strData = strData & "&entry.1751698587=" & Replace(EndDate, "#", "No.")
    strData = strData & "&entry.1576486540=" & Replace(Nor, "#", "No.")
    strData = strData & "&entry.39225992=" & Replace(Erd, "#", "No.")
    strData = strData & "&entry.1764030453=" & Replace(MaxEndDate, "#", "No.")
    strData = strData & "&entry.2100539981=" & Replace(Remarks, "#", "No.")
    strData = strData & "&entry.1179689168=" & Replace(Notes, "#", "No.")
    strData = strData & "&entry.1268903543=" & Replace(ExtensionDuration, "#", "No.")
    strData = strData & "&entry.1058668284=" & Replace(EstimatedSpend, "#", "No.")
    strData = strData & "&entry.2088119780=" & Replace(ContractLinked, "#", "No.")
    strData = strData & "&entry.1445962386=" & Replace(Files, "#", "No.")
    strData = strData & "&entry.1078024577=" & Replace(Priority, "#", "No.")
    strData = strData & "&entry.1474710638=" & Replace(NextRenDate, "#", "No.")
    strData = strData & "&entry.1712999813=" & Replace(DaysLforRen, "#", "No.")
    strData = strData & "&entry.1143322447=" & Replace(CurrRenPeriod, "#", "No.")
    strData = strData & "&entry.2025005756=" & Replace(CLMSReqNumber, "#", "No.")
    strData = strData & "&entry.1560215212=" & Replace(SupplierABNumber, "#", "No.")
    strData = strData & "&entry.479798142=" & Replace(SharepointNumber, "#", "No.")
    strData = strData & "&entry.1971403833=" & Replace(RecDateEntered, "#", "No.")
    strData = strData & "&entry.1662468430=" & Replace(ContractDateEntered, "#", "No.")
    strData = strData & "&entry.1416211103=" & Replace(RenewContract, "#", "No.")
    
    If Unique_ID = "" Then
        ws.Range("AR" & i).Value = pk_id
        strData = strData & "&entry.1678261984=" & pk_id
    Else
        strData = strData & "&entry.1678261984=" & Unique_ID
    End If
    strData = strData & "&entry.212087561=" & IIf(DeleteContract = "", "No", DeleteContract)
    
    strData = strData & "&entry.835752841=No"
    
    'strData = Replace(strData, "% ", "No.")
    
        'Debug.Print i, ContractNumber
        finalURL = URL & strData
        'finalBackupURL = backupURL & strData
        http.Open "POST", finalURL, False
        
        'http2.Open "POST", finalBackupURL, False
        
        'http2.send
        http.send
        
        'Debug.Print http.statusText
            
            If DeleteContract = "Yes" Then
            
                deletedFlag = True
                ReDim Preserve arrRowsTobeDeleted(intArrsize)
                arrRowsTobeDeleted(intArrsize) = i
                intArrsize = intArrsize + 1
            
            End If
            
            
            If http.statusText = "OK" Then
                
                ws.Range("AT" & i).Value = "Synced"
                If Unique_ID = "" Then
                    
                    pk_id = pk_id + 1
                    Sheet8.Range("FA1").Value = pk_id
                    Sheet8.Range("FF1").Value = Now
                
                End If
                
            Else
            
            Debug.Print strData
            Debug.Print i & vbNewLine
            Debug.Print http.statusText, http2.statusText
            
            updateLog ThisWorkbook, "request text: " & http.statusText, "SyncData: Failed"
            
            End If
            
    End If
Next i

Call deleteRows
'MsgBox "Data Synced!"
    
End Sub

Function deleteRows()


    If deletedFlag = True Then
        For i = UBound(arrRowsTobeDeleted) To LBound(arrRowsTobeDeleted) Step -1
        
            Sheet8.Range("A" & arrRowsTobeDeleted(i)).Resize(1, Sheet8.ListObjects(1).ListColumns.count).Cells.Delete
        
        Next i
    End If
    
    deletedFlag = False
    
End Function

