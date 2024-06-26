Attribute VB_Name = "FrontEndCodes"
Option Explicit
'This Page Holds Code to Execute the FrontEnd looks of each page

'This Visual Basic code enhances the appearance and functionality of a worksheet named 'ws' within a _
Microsoft Excel spreadsheet. It merges certain cell ranges to create larger cells, adjusts row heights _
and column widths, hides specific rows, unlocks certain cells, and sets up a named range for data entry
'The code also adds a background picture to the worksheet and turns off the display of gridlines. _
It restores the default Excel settings upon completion.

Sub renameShapes()
    
    Dim sh As Worksheet
    
    unProtectWorksheet
    On Error Resume Next
    For Each sh In ThisWorkbook.Worksheets
            
        If InStr(1, sh.Range("A1").Value, "Nav", vbTextCompare) > 0 Then
            
            sh.Shapes("Info_profileName").TextFrame.Characters.Text = Sheet12.Range("pName").Value
            
            sh.Shapes("Heading_AppName").TextFrame.Characters.Text = Replace(ThisWorkbook.name, ".xlsm", "")
        
        End If
        
    Next sh
    
    protectWorksheet

End Sub

Sub PageDesign(ws As Worksheet)
    Dim Settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    Settings.TurnOn
    Settings.TurnOff
    
    unprotc ws
    
    Dim i As Long
    
    'Sheet1.Range("A:C").Merge
    On Error Resume Next
    
    For i = 1 To 12 Step 2
        
        
        ws.Range("A" & i & ":C" & i + 1).Merge
    
    Next i
    
    For i = 13 To 100
        
        ws.Range("A" & i & ":C" & i).Merge
    
    Next i
    
    ws.Range("A1:C12").RowHeight = 15
    
    ws.Range("A13:C100").RowHeight = 25
    
    ws.Range("A:C").ColumnWidth = 10
    
    ws.Range("D:Z").ColumnWidth = 30
    
    If ws.name <> Sheet1.name Then
        
        Range("16:16").rows.Hidden = True
        
        Range("D16:T17").Locked = False
    
    End If
    
    ws.Range("D11:E12").Merge
    
    ws.Range("D11:E12").name = "ProfileName"
    
    ws.Range("ProfileName").HorizontalAlignment = xlHAlignLeft
    
    'ws.SetBackgroundPicture Filename:=Sheet1.Shapes("Info_Root_Dir").TextFrame.Characters.Text & "\User Data\Bg.jpg"
    
    ActiveWindow.DisplayGridlines = False
    
    protc ws
    
    Settings.Restore

End Sub

'Procedure to Set the theme of the page like button formats and cell background and table design Etc..
Sub Theme()
    'To change Theme goto sheet21 and change the format of the Shapes in the page and it will be reflected in all the pages
    'Four types of elements in a page
    '1) Button = 'Btn'
    '2) Active Button = 'Active'
    '3) Heading = 'Heading'
    '4) Information = 'Info'
    
    Dim shp As Shape, sh As Shape
    'Debug.print Sheet21.Shapes("Button").Fill.GradientColorType
    
    unprotc
    On Error Resume Next
    For Each shp In ActiveSheet.Shapes
        
        
                
        If InStr(1, shp.name, "Active", vbTextCompare) > 0 Then
            
            Sheet21.Shapes("ActiveButton").PickUp

            shp.Apply
        
        ElseIf Left(shp.name, 3) = "Btn" Then
                
            Sheet21.Shapes("Button").PickUp
            
            shp.Apply

        
        End If
        
        If Left(shp.name, 7) = "Heading" Then
            
            Sheet21.Shapes("Heading").PickUp
            
            shp.Apply
        
        End If
        
        If Left(shp.name, 4) = "Info" Then
            
            Sheet21.Shapes("Info").PickUp
            
            shp.Apply
        
        End If
    
    Next shp
    
    'Sheet1.Shapes("Info_Root_Dir").TextFrame.Characters.Font.Size = 9
    
    protc
        
End Sub


Sub changeTitle()
Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
    
        If ws.Range("A1").Value = "NavTo" Then
        
            unprotc ws
            
                ws.Shapes("Heading_AppName").TextFrame.Characters.Text = ThisWorkbook.name
                
            protc ws
        
        End If
    
    Next ws
    
End Sub


Sub ApplyfrontEnd()

    Dim Settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    Dim Timer As New TimerCls
    
    Settings.TurnOn
    
    Settings.TurnOff

        PageDesign ActiveSheet
        
        AddShapes ActiveSheet
        
        Theme
    
    Settings.Restore

End Sub

Sub AddShapes(newSheet As Worksheet)
'This code creates the shapes in the Home page and assigns them the function they need to do onAction
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim PageName As Shape, applicationName As Shape, rootDir As Shape, browseFolder As Shape
    Dim lft As Double, tp As Double, wdth As Double, hight As Double, shp As Shape

    'Set the worksheet object
    Set ws = newSheet
    
    unprotc ws
    
    For Each shp In ws.Shapes
    
        If Left(shp.name, 3) = "Btn" Then
        
            shp.Delete
        
        End If
    
    Next shp
    
    On Error Resume Next
    
    'if exists delete existing shape and create a newshape and format
    'AppName
        
        Set rng = ws.Range("A1:Z2")
        
        lft = 0
        
        tp = 0
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete existing shape
        
        ws.Shapes("Heading_AppName").Delete
        
        'Create a new shape
        
        Set applicationName = ws.Shapes.AddShape(msoShapeRectangle, lft, tp, wdth, hight)
        
        applicationName.name = "Heading_AppName"
        
        'Adjust parameters
        
        applicationName.TextFrame.VerticalAlignment = xlVAlignCenter
        
        applicationName.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        applicationName.TextFrame.Characters.Font.Bold = True
        
        applicationName.Line.Visible = msoFalse
        
        applicationName.TextFrame.Characters.Text = Replace(ThisWorkbook.name, ".xlsm", "")
    
    'Add a navigation button to pages which should be visible in this application
    
    'Tag the pages as VISIBLE in range("A1") if the page should be seen by the user
    
    'NavigateTab
        Dim pg As Worksheet
        Dim navShape As Shape
        
        Set rng = ws.Range("A3:C4")
        
        lft = 0
        
        tp = rng.Top
        
        wdth = rng.Width - 2
        
        hight = rng.Height
        'Delete
        ws.Shapes("Heading_NavigateTab").Delete
        'Create
        Set navShape = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, 0, tp, wdth, hight)
        
        navShape.name = "Heading_NavigateTab"
        'Edit
        navShape.TextFrame.VerticalAlignment = xlVAlignCenter
        
        navShape.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        navShape.Line.Visible = msoFalse
        
        navShape.TextFrame.Characters.Text = "Navigation Menu"
        
        'Navigation
        Dim pageShape As Shape, pageShapeGroup As Shape, count As Integer
        
        Dim groups() As String, grp As String
        
        count = 1
        
        lft = navShape.Left
        
        tp = navShape.Top
        
        wdth = navShape.Width
        
        hight = navShape.Height
        'Delete

        
        
        'Create
         For Each pg In ThisWorkbook.Worksheets
             
             If LCase(pg.Range("A1")) = "navto" Then
                 
                 Set pageShape = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp + _
                                     hight * count, wdth, hight)
                 If pg.name = ws.name Then
                     
                     pageShape.name = "Btn_Active"
                 
                 Else

                     pageShape.name = "Btn_" & (LCase(Trim(Replace(pg.name, " ", ""))))
                 
                 End If
                 'Edit
                 pageShape.TextFrame.VerticalAlignment = xlVAlignCenter
                 
                 pageShape.TextFrame.HorizontalAlignment = xlHAlignLeft
                 

                 
                 pageShape.Line.Visible = msoFalse
                 
                 pageShape.TextFrame.Characters.Text = pg.name
                 
                 pageShape.OnAction = "navtoSheet" 'Replace(pg.Name, " ", "_")
                 
                 pageShape.Select Replace:=False
                 
                 count = count + 1
            
            ElseIf pg.name = Sheet16.name Then
                 
                 Set pageShape = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp + _
                                     hight * count, wdth, hight)
                 If pg.name = ws.name Then
                 

                     pageShape.name = "Btn_Active"
                 
                 Else

                     pageShape.name = "Btn_" & (LCase(Trim(Replace(pg.name, " ", ""))))
                 
                 End If
                 'Edit
                 pageShape.TextFrame.VerticalAlignment = xlVAlignCenter
                 
                 pageShape.TextFrame.HorizontalAlignment = xlHAlignLeft
                 

                 
                 pageShape.Line.Visible = msoFalse
                 
                 pageShape.TextFrame.Characters.Text = pg.name
                 
                 pageShape.OnAction = "navtoSheet" 'Replace(pg.Name, " ", "_")
                 
                 'pageShape.Select Replace:=False
                 
                 count = count + 1
    
             ElseIf LCase(pg.Range("A1")) = "nav_to" Then
                 
                 Set pageShape = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp + _
                                     hight * count, wdth, hight)
                 If pg.name = ws.name Then

                     pageShape.name = "Btn_Active"
                 
                 Else

                     pageShape.name = "Btn_" & (LCase(Trim(Replace(pg.name, " ", ""))))
                 
                 End If
                 'Edit
                 pageShape.TextFrame.VerticalAlignment = xlVAlignCenter
                 
                 pageShape.TextFrame.HorizontalAlignment = xlHAlignLeft
                 

                 
                 pageShape.Line.Visible = msoFalse
                 
                 pageShape.TextFrame.Characters.Text = pg.name
                 
                 pageShape.OnAction = "" 'Replace(pg.Name, " ", "_")
                 
                 'pageShape.Select Replace:=False
                 
                 count = count + 1
             
             End If
             
         Next pg
        
        ws.Range("D3").Select
        
    'Only run for home page
    
    If ws.name = Sheet1.name Then
    
    'Sign-In
        Dim SignIn As Shape
        
        Set rng = ws.Range("D3:D4")
        lft = navShape.Width
        tp = rng.Top
        wdth = rng.Width
        hight = rng.Height
        
    'Create
        Set SignIn = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        SignIn.name = "Btn_SignInOut"
        
    'Edit
        
        SignIn.TextFrame.VerticalAlignment = xlVAlignCenter
        SignIn.TextFrame.HorizontalAlignment = xlHAlignCenter

        SignIn.Line.Visible = msoFalse
        
        If Sheet12.Range("pName").Value = "" Then
            
            SignIn.TextFrame.Characters.Text = "Sign In"
            SignIn.OnAction = "SignInCode"
        
        Else
            
            SignIn.TextFrame.Characters.Text = "Sign Out"
            SignIn.OnAction = "SignOutCode"
        
        End If
    
    'staffManagement
        
        Dim StaffMgmt As Shape
        
        Set rng = Sheet1.Range("E3:E4")
        lft = SignIn.Left + SignIn.Width
        tp = rng.Top
        wdth = rng.Width
        hight = rng.Height
        
    'Create
        
        Set StaffMgmt = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        StaffMgmt.name = "Btn_StaffMgmt"
        
    'edit
        
        StaffMgmt.TextFrame.VerticalAlignment = xlVAlignCenter
        StaffMgmt.TextFrame.HorizontalAlignment = xlHAlignCenter

        StaffMgmt.Line.Visible = msoFalse
        StaffMgmt.TextFrame.Characters.Text = "Staff Management"
        
        If Range("Security").Value = "Admin" Then
            
            StaffMgmt.Visible = msoCTrue
            StaffMgmt.OnAction = "openStaffMgmt"
        
        Else
            
            StaffMgmt.Visible = msoFalse
        
        End If
      
    'Portfolio Management
    
        Dim portfolioManagement As Shape
        
        Set rng = Sheet1.Range("E3:E4")
        lft = StaffMgmt.Left + StaffMgmt.Width
        tp = rng.Top
        wdth = rng.Width
        hight = rng.Height

    'Create
        
        Set portfolioManagement = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        portfolioManagement.name = "Btn_PortfMgmt"
        
    'edit
        
        portfolioManagement.TextFrame.VerticalAlignment = xlVAlignCenter
        
        portfolioManagement.TextFrame.HorizontalAlignment = xlHAlignCenter
        

        
        portfolioManagement.Line.Visible = msoFalse
        
        portfolioManagement.TextFrame.Characters.Text = "Portfolio Management"
        
        If Range("Security").Value = "Admin" Then
            
            portfolioManagement.Visible = msoCTrue
            
            portfolioManagement.OnAction = "openPortfolioMgmt"
        
        Else
            
            portfolioManagement.Visible = False
        
        End If
    
    'Contracts Field Management
    
        Dim fieldAccess As Shape
        
        Set rng = Sheet1.Range("E3:E4")
        
        lft = portfolioManagement.Left + portfolioManagement.Width
        
        tp = rng.Top
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
        

        
        'Create
        
        Set fieldAccess = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        fieldAccess.name = "Btn_FieldAccess"
        
        'edit
        
        fieldAccess.TextFrame.VerticalAlignment = xlVAlignCenter
        
        fieldAccess.TextFrame.HorizontalAlignment = xlHAlignCenter
        

        
        fieldAccess.Line.Visible = msoFalse
        
        fieldAccess.TextFrame.Characters.Text = "Fields Management"
        
        If Range("Security").Value = "Admin" Then
            
            fieldAccess.Visible = msoCTrue
            
            fieldAccess.OnAction = "fieldsManagement"
        
        Else
            
            fieldAccess.Visible = False
        
        End If
        
    End If
        
    'PageHead
        Set rng = ws.Range("D9:E10")
        
        lft = navShape.Width
        
        tp = rng.Top
        
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
        
        ws.Shapes("Heading_PageHead").Delete
        
        'Create
        
        Set PageName = ws.Shapes.AddShape(msoShapeRectangle, lft, tp, wdth, hight)
        
        PageName.name = "Heading_PageHead"
        
        'Edit
        
        PageName.TextFrame.Characters.Text = ws.name
        
        PageName.TextFrame.VerticalAlignment = xlVAlignCenter
        
        PageName.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        PageName.TextFrame.Characters.Font.Bold = True
        
        PageName.Line.Visible = msoFalse
        
    
    'Create ProfileName Shape
                        
        Dim profileNm As Shape
        
        Set rng = ws.Range("ProfileName")
        
        lft = PageName.Left
        
        tp = PageName.Top + PageName.Height
        
        wdth = PageName.Width
        
        hight = PageName.Height
        
        'Delete
         
         ws.Shapes("Info_profileName").Delete
         
         'Create
        
        Set profileNm = ws.Shapes.AddShape(msoShapeRectangle, lft, tp, wdth, hight)
        
        profileNm.name = "Info_profileName"
        
        'Edit
        
        profileNm.Line.Visible = msoFalse
        
        profileNm.TextFrame.VerticalAlignment = xlVAlignCenter
        
        profileNm.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        profileNm.TextFrame.Characters.Text = Range("pName").Value
        
        'profileNm.Visible = msoTrue
                
    If ws.name <> Sheet1.name Then
    
    'Apply Filter
        Dim ApplyFilter As Shape
        
        Set rng = ws.Range("F9:F10")
        
        lft = PageName.Left + PageName.Width
        
        tp = rng.Top
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
         

         'Create
        
        Set ApplyFilter = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        ApplyFilter.name = "Btn_Filters"
        
        'Edit
        
        ApplyFilter.Line.Visible = msoFalse
        
        ApplyFilter.TextFrame.VerticalAlignment = xlVAlignCenter
        
        ApplyFilter.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        ApplyFilter.TextFrame.Characters.Text = "Apply Filters"
        
        ApplyFilter.OnAction = "applyAdvFilt"
    
    'ClearFilters
        
        Dim ClrFilters As Shape
        
        Set rng = ws.Range("F9:F10")
        
        lft = ApplyFilter.Left + ApplyFilter.Width
        
        tp = rng.Top
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
         

         
         'Create
        
        Set ClrFilters = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        ClrFilters.name = "Btn_clrFilters"
        
        'Edit
        
        ClrFilters.Line.Visible = msoFalse
        
        ClrFilters.TextFrame.VerticalAlignment = xlVAlignCenter
        
        ClrFilters.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        ClrFilters.TextFrame.Characters.Text = "Clear Filters"
        
        ClrFilters.OnAction = "clearFilters"
    
    'Edit Contract
        
        Dim editContr As Shape
        
        Set rng = ws.Range("F9:F10")
        
        lft = ClrFilters.Left + ClrFilters.Width
        
        tp = rng.Top
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
         

         
         'Create
        
        Set editContr = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        editContr.name = "Btn_editContr"
        
        'Edit
        
        editContr.Line.Visible = msoFalse
        
        editContr.TextFrame.VerticalAlignment = xlVAlignCenter
        
        editContr.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        editContr.TextFrame.Characters.Text = "Edit Contract"
        
        editContr.OnAction = "openContractEdit"
        
    'AddNewContract
        
        Dim newContract As Shape
        
        Set rng = ws.Range("I9:I10")
        
        lft = editContr.Left + editContr.Width
        
        tp = rng.Top
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
         

         
         'Create
        
        Set newContract = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        newContract.name = "Btn_newContr"
        
        'Edit
        
        newContract.Line.Visible = msoFalse
        
        newContract.TextFrame.VerticalAlignment = xlVAlignCenter
        
        newContract.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        newContract.TextFrame.Characters.Text = "Add New Contract"
        
        newContract.OnAction = "addNewContract"
    
    
    'CreateReport
        
        Dim CreateReport As Shape
        
        Set rng = ws.Range("J9:J10")
        
        lft = newContract.Left + newContract.Width
        
        tp = rng.Top
        
        wdth = newContract.Width
        
        hight = rng.Height
        
        'Delete
         

         
         'Create
        
        Set CreateReport = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        CreateReport.name = "Btn_CreateReport"
        
        'Edit
        
        CreateReport.Line.Visible = msoFalse
        
        CreateReport.TextFrame.VerticalAlignment = xlVAlignCenter
        
        CreateReport.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        CreateReport.TextFrame.Characters.Text = "Create Report"
        
        CreateReport.OnAction = "CreateReport"
    
    
    'SearchPage
        
        Dim SearchPage As Shape
        
        Set rng = ws.Range("J9:J10")
        
        lft = CreateReport.Left + CreateReport.Width
        
        tp = rng.Top
        
        wdth = CreateReport.Width
        
        hight = rng.Height
        
        'Delete
         

         
         'Create
        
        Set SearchPage = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        SearchPage.name = "Btn_SearchPage"
        
        'Edit
        
        SearchPage.Line.Visible = msoFalse
        
        SearchPage.TextFrame.VerticalAlignment = xlVAlignCenter
        
        SearchPage.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        SearchPage.TextFrame.Characters.Text = "Search Page"
        
        SearchPage.OnAction = "FindCode"
    
    Else
    
        Dim RefreshDashboard As Shape
        
        Set rng = ws.Range("F9:F10")
        
        lft = PageName.Left + PageName.Width
        
        tp = rng.Top
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
         

         'Create
        
        Set RefreshDashboard = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        RefreshDashboard.name = "Btn_RefDash"
        
        'Edit
        
        RefreshDashboard.Line.Visible = msoFalse
        
        RefreshDashboard.TextFrame.VerticalAlignment = xlVAlignCenter
        
        RefreshDashboard.TextFrame.HorizontalAlignment = xlHAlignLeft
        

        
        RefreshDashboard.TextFrame.Characters.Text = "Refresh Dashboard"
        
        RefreshDashboard.OnAction = "RefDashboard"
    
    End If
    
    
    If ws.name = Sheet16.name Then
    
    'RefreshPage
        
        Dim RefreshPage As Shape
        
        Set rng = ws.Range("F11:F12")
        
        lft = profileNm.Left + profileNm.Width
        
        tp = profileNm.Top
        
        wdth = rng.Width
        
        hight = rng.Height
        
        'Delete
         

         
         'Create
        
        Set RefreshPage = ws.Shapes.AddShape(msoShapeRound2DiagRectangle, lft, tp, wdth, hight)
        
        RefreshPage.name = "Btn_RefPage"
        
        'Edit
        
        RefreshPage.Line.Visible = msoFalse
        
        RefreshPage.TextFrame.VerticalAlignment = xlVAlignCenter
        
        RefreshPage.TextFrame.HorizontalAlignment = xlHAlignLeft
        
        RefreshPage.TextFrame.Characters.Text = "Refresh All Data"
            
        RefreshPage.OnAction = "refreshAllContracts"
    
    End If
    
    For Each shp In ws.Shapes
    
        If Left(shp.name, 3) = "Btn" Then

            shp.Placement = xlFreeFloating
            
        End If
        
    Next shp
 protc ws
 
End Sub

Sub textSize(ws As Worksheet)

Dim shp As Shape
unprotc ws
protc ws
End Sub
