Attribute VB_Name = "DownloadData"
'https://script.google.com/macros/s/AKfycbzESMb0FcqAkkmu8TXXnkeL1EAR1ZBMY5V3j8s4GB4mCsHcDi5Lr3HfUVhb7AxWU1JW3g/exec
Sub getUpdatedData()

    Dim settings As New ExclClsSettings
        
    settings.TurnOn
    settings.TurnOff
    

    Dim Httpreq As Object
    Dim URL As String, response As String, arr() As String
    
    URL = "https://script.google.com/macros/s/AKfycbzESMb0FcqAkkmu8TXXnkeL1EAR1ZBMY5V3j8s4GB4mCsHcDi5Lr3HfUVhb7AxWU1JW3g/exec"
    Set Httpreq = CreateObject("MSXML2.ServerXMLHTTP")
    
    With Httpreq
        .Open "GET", URL, False
        .send
    End With
    
    Do Until Httpreq.readyState = 4: Loop
    
    response = Httpreq.responseText
    
    'Debug.Print response
    DoEvents
    Call HTMLtoRange(response, Sheet8)
    
    settings.Restore
End Sub


Sub getResponseFromSheets(URL As String, targetWs As Worksheet)

    Dim Httpreq As Object
    Dim response As String, arr() As String
    
    Set Httpreq = CreateObject("MSXML2.ServerXMLHTTP")
    
    With Httpreq
        .Open "GET", URL, False
        .send
    End With
    
    Do Until Httpreq.readyState = 4: Loop
    
    response = Httpreq.responseText
    
    'Debug.Print response
    DoEvents
    Call HTMLtoRange(response, targetWs)

End Sub



Sub HTMLtoRange(Data, sheet As Worksheet)

    Dim HTML As Object
    Dim myDate As New dateConv
    Dim arr() As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Set HTML = CreateObject("htmlFile")
    
    Set ws = sheet
    
    HTML.body.innerHTML = Data
    
   ' Debug.Print Data
    
    r = 0
     ws.ListObjects(1).DataBodyRange.ClearContents
    
    row = (Len(Data) - Len(Replace(Data, "<tr>", ""))) / 4
    
    col = (Len(Data) - Len(Replace(Data, "<td>", ""))) / (row * 4)
    
    ReDim arr(1 To row, 1 To col + 2)
    
    
    For Each tr In HTML.getElementsByTagName("tr")

        r = r + 1
        c = 0
        For Each td In tr.getElementsByTagName("td")

            c = c + 1

            If InStr(1, td.innerText, "GMT-", vbTextCompare) > 0 Then
                arr(r, c) = myDate.ConvertDates(td.innerText)
            Else
                arr(r, c) = td.innerText

            End If
        Next td
        If r = 1 Then
            ws.Cells(r, ws.Range("A1").CurrentRegion.Columns.count - 1).Value = "To_Be_Deleted"
            ws.Cells(r, ws.Range("A1").CurrentRegion.Columns.count).Value = "SyncStatus"
        Else
            arr(r, col + 1) = "No"
            arr(r, col + 2) = "Synced"
        End If
    Next tr

    ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    
        ws.Cells(1, ws.Range("A1").CurrentRegion.Columns.count - 1).Value = "To_Be_Deleted"
        ws.Cells(1, ws.Range("A1").CurrentRegion.Columns.count).Value = "SyncStatus"
    
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    
End Sub

