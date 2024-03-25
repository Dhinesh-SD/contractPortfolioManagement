Attribute VB_Name = "StaffDataSync"
Option Explicit
Dim arrRowsTobeDeleted()
Dim intArrsize
Dim deletedFlag As Boolean
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

'https://docs.google.com/forms/d/e/1FAIpQLScqWGLDRfOBekHe5m74ybswBOLtcM5fZ8gGQqEeiHzJFm-F0g/viewform?usp=pp_url&entry.1504783255=SID&entry.765243750=FN&entry.2117007185=LN&entry.135352479=FUN&entry.490415601=EID&entry.780068068=POS&entry.1452816150=BID&entry.1741837361=SEC&entry.1277509256=FIELD&entry.1577364834=TBD&entry.862266060=SYNCED&entry.1449018300=Logged_Out
Sub sync()
SyncstaffData_FromGsheets
End Sub

Sub SyncStaffData_ToGsheets()
    
    Dim http As Object
    Dim URL As String
    Dim Primary_Key As String
    Dim pk_id As Long
    
    Dim i As Long
    Dim strData As String
    Dim finalURL As String
    Dim ws As Worksheet
    
    Dim staffId As String, F_Name As String, L_Name As String, Full_Name As String, Email As String
    Dim position As String, BuyerId As String, Security As String, pass As String, tobeDel As String
    Dim syncStatus As String, loginStatus As String
    
    Set ws = Sheet6
    URL = "https://docs.google.com/forms/d/e/1FAIpQLScqWGLDRfOBekHe5m74ybswBOLtcM5fZ8gGQqEeiHzJFm-F0g/formResponse?"
    Set http = CreateObject("MSXML2.ServerXMLHTTP")

    pk_id = 0
    deletedFlag = False
    intArrsize = 0
    
    pk_id = Application.WorksheetFunction.Max(ws.Range("A2:A" & ws.Cells(2, 1).End(xlDown).row)) + 1
    
    For i = 2 To ws.Range("A1").End(xlDown).row
    

        staffId = ws.Cells(i, "A").Value
        F_Name = ws.Cells(i, "B").Value
        L_Name = ws.Cells(i, "C").Value
        Full_Name = ws.Cells(i, "D").Value
        Email = ws.Cells(i, "E").Value
        position = ws.Cells(i, "F").Value
        BuyerId = ws.Cells(i, "G").Value
        Security = ws.Cells(i, "H").Value
        pass = ws.Cells(i, "I").Value
        tobeDel = ws.Cells(i, "K").Value
        syncStatus = ws.Cells(i, "L").Value
        loginStatus = ws.Cells(i, "J").Value

        If staffId = "" Then
        
            ws.Range("A" & i).Value = pk_id
            
            strData = "&entry.1504783255==" & pk_id
            
        Else
        
            strData = "&entry.1504783255=" & staffId
        
        End If
        
'&entry.1504783255=StaffID
'&entry.765243750=FN
'&entry.2117007185=LN
'&entry.135352479=FUN
'&entry.490415601=EID
'&entry.780068068=POS
'&entry.1452816150=BID
'&entry.1741837361=SEC
'&entry.1277509256=FIE
'&entry.1577364834=TBD
'&entry.862266060=SYNC

        strData = strData & "&entry.765243750=" & F_Name
        strData = strData & "&entry.2117007185=" & L_Name
        strData = strData & "&entry.135352479=" & Full_Name
        strData = strData & "&entry.490415601=" & Email
        strData = strData & "&entry.780068068=" & position
        strData = strData & "&entry.1452816150=" & BuyerId
        strData = strData & "&entry.1741837361=" & Security
        strData = strData & "&entry.1277509256=" & pass
        strData = strData & "&entry.1577364834=" & IIf(tobeDel = "", "No", tobeDel)
        strData = strData & "&entry.862266060=" & "No"
        strData = strData & "&entry.1449018300=" & loginStatus
        
        
        If syncStatus = "" Then
            'Debug.Print strData
            finalURL = URL & strData
            http.Open "POST", finalURL, False
            http.send

            If tobeDel = "Yes" Then
                
                    deletedFlag = True
                    ReDim Preserve arrRowsTobeDeleted(intArrsize)
                    arrRowsTobeDeleted(intArrsize) = i
                    intArrsize = intArrsize + 1
                
            End If
                
                
            If http.statusText = "OK" Then
                
                unprotc ws
                ws.Range("L" & i).Value = "Synced"
                Sheet6.Range("T1").Value = Now
                protc ws
                
            Else
            
                MsgBox http.statusText
                Debug.Print finalURL
                
            End If
        
        End If
        
    Next i
    
deleteRows2
    
End Sub




Sub SyncstaffData_FromGsheets(Optional typ As String)
    
    Dim URL As String
    
    Dim seconds As Integer
    
    seconds = 120
    
    If DateDiff("s", Sheet6.Range("T1").Value, Now) < seconds Or typ <> "Mandatory" Then
    
        Exit Sub
    
    End If
    
    unprotc Sheet6
    
    
    URL = "https://script.google.com/macros/s/AKfycbzUmapGRuJ5Y1OFfkz6xcp1OyFvW7-p2LNr-U8GfWGs_zPeThuqyzGNmwV-XOY7QeCU/exec"

    getResponseFromSheets URL, Sheet6
        
    Sheet6.Range("T1").Value = Now
        
    protc Sheet6
End Sub



Function deleteRows2()

unprotc Sheet6
    Dim i As Long
    If deletedFlag = True Then
        For i = UBound(arrRowsTobeDeleted) To LBound(arrRowsTobeDeleted) Step -1
        
            Sheet6.Range("A" & arrRowsTobeDeleted(i)).Resize(1, Sheet6.ListObjects(1).ListColumns.count).Cells.Delete
        
        Next i
    End If
    
    deletedFlag = False
protc Sheet6
End Function
