Attribute VB_Name = "NavigationCodes"
Option Explicit
Sub Navto(wsheet As Worksheet)

    Dim Settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    Settings.TurnOff
    
    Dim sh As Worksheet
    
    For Each sh In ThisWorkbook.Worksheets
        
        If sh.name = wsheet.name Then
            
            wsheet.Visible = True
            Application.enableEvents = True
            
            ThisWorkbook.Worksheets(wsheet.name).Select
            
            Application.enableEvents = False
            Exit For
        
        End If
    
    Next sh
    
    For Each sh In ThisWorkbook.Worksheets
        
        If sh.name <> wsheet.name Then sh.Visible = False
    
    Next sh
    
    'ApplyfrontEnd
    
    If ActiveSheet.name <> Sheet1.name Then
        
        If ActiveSheet.Range("A3").Value = True Then
            
            applyAdvFilt
            
            ActiveSheet.Range("A3").Value = False
        
        End If
        
        If ActiveSheet.Shapes("Info_profileName").TextFrame.Characters.Text <> Sheet12.Range("pName").Value Then
            
        unprotc ActiveSheet
        
            ActiveSheet.Shapes("Info_profileName").TextFrame.Characters.Text = Sheet12.Range("pName").Value
            ActiveSheet.Shapes("Heading_AppName").TextFrame.Characters.Text = Replace(ThisWorkbook.name, ".xlsm", "")
        
        protc ActiveSheet
        End If
        Range("A18").Select
                
        ActiveWindow.FreezePanes = True
    
    End If
    
    ActiveWindow.Zoom = 90
    
    Settings.Restore
    
End Sub


Sub navtoSheet()

    Dim ws As Worksheet
    Dim wsName As String
    Dim name As String
    
    name = Application.Caller
    
    If InStr(1, Application.Caller, "replace", vbTextCompare) > 0 Then name = "Btn_Procurements(ReplaceExisting)"
    
    wsName = ActiveSheet.Shapes(name).TextFrame.Characters.Text
    
        If ThisWorkbook.Worksheets(wsName).Range("A1").Value <> "NavTo" Then Exit Sub
        
        Navto ThisWorkbook.Worksheets(wsName)

End Sub

