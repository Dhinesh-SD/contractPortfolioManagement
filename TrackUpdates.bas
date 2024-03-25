Attribute VB_Name = "TrackUpdates"
Option Explicit

'https://docs.google.com/forms/d/e/1FAIpQLScatSzReYvqp-DK4lfmzMi02BVHP9c0PvXsW-LG-3jS5yEWtw/viewform?usp=pp_url
'&entry.1157438081=aaacsc&entry.823026421=ADWFWFS&entry.2073121590=FWFVCSC
Sub updateLog(wb As Workbook, Remarks As String, Procedure As String)

    Dim settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    Dim URL As String
    
    Dim strData As String
    
    Dim updtTbl As ListObject
    
    Dim ws As Worksheet
    
    Dim UpdatedBy As String
    
    Dim http As Object
    
    Dim finalURL As String
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    URL = "https://docs.google.com/forms/d/e/1FAIpQLScatSzReYvqp-DK4lfmzMi02BVHP9c0PvXsW-LG-3jS5yEWtw/formResponse?"
    
    UpdatedBy = ThisWorkbook.Worksheets("Profile Information").Range("B5").Value
    
    strData = "&entry.1157438081=" & UpdatedBy
    strData = strData & "&entry.823026421=" & Remarks
    strData = strData & "&entry.2073121590=" & Procedure
    
    finalURL = URL & strData
    
    http.Open "POST", finalURL, False
    http.send
    
    settings.Restore
    
    
End Sub


Sub unlck()
    Dim sh As Worksheet
    Dim settings As New ExclClsSettings
    
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff

    For Each sh In ThisWorkbook.Worksheets
        
        If sh.Range("A1").Value = "NavTo" And sh.name <> Sheet1.name Then
            
            unprotc sh
                
                sh.Range("A3:C4").Locked = False
                
                sh.Range("A3").Value = False
            
            protc sh
        
        End If
    
    Next sh
    
    settings.Restore
    
End Sub

