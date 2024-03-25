VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditContracts2 
   Caption         =   "Edit Contracts Information"
   ClientHeight    =   12432
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   16872
   OleObjectBlob   =   "EditContracts2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditContracts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'
'Private Sub setProperties(cntrl1 As Control, cntrl2 As Control)
'
'    cntrl1.BackColor = cntrl2.BackColor
'    cntrl1.ForeColor = cntrl2.ForeColor
'    cntrl1.BackStyle = cntrl2.BackStyle
'    cntrl1.BorderStyle = cntrl2.BorderStyle
'    cntrl1.BorderColor = cntrl2.BorderColor
'
'End Sub
'
'

'
'

'
'
'
'Private Sub highlight(cntrl As Control)
'
'If Not cntrl.BackColor = ThemeUf.SampleActive.BackColor Then
'
'    setProperties cntrl, ThemeUf.SampleActive
'
'End If
'
'End Sub
'
'
'Private Sub Exit_highlight(cntrl As Control)
'
'If Not cntrl.BackColor = ThemeUf.SampleInactive.BackColor Then
'
'    setProperties cntrl, ThemeUf.SampleInactive
'
'End If
'
'End Sub
'
'Private Sub Onclick(cntrl As Control)
'
'    setProperties cntrl, ThemeUf.onClick_Btn
'
'End Sub



Private Sub checkHighlight()
        Select Case (Me.MultiPage1.Value)
    
        Case 0
            Exit_highlight Me.Btn_DeviationDetails
            Exit_highlight Me.Btn_Amendments
            Exit_highlight Me.Btn_RenewalInfo
            Exit_highlight Me.Btn_Status
        Case 1
            Exit_highlight Me.Btn_GeneralInfo
            Exit_highlight Me.Btn_Amendments
            Exit_highlight Me.Btn_RenewalInfo
            Exit_highlight Me.Btn_Status
        Case 2
            Exit_highlight Me.Btn_DeviationDetails
            Exit_highlight Me.Btn_RenewalInfo
            Exit_highlight Me.Btn_GeneralInfo
            Exit_highlight Me.Btn_Status
        Case 3
            Exit_highlight Me.Btn_DeviationDetails
            Exit_highlight Me.Btn_Amendments
            Exit_highlight Me.Btn_Status
            Exit_highlight Me.Btn_GeneralInfo
        Case 4
            Exit_highlight Me.Btn_DeviationDetails
            Exit_highlight Me.Btn_Amendments
            Exit_highlight Me.Btn_RenewalInfo
            Exit_highlight Me.Btn_GeneralInfo
    End Select
End Sub



Private Sub ActBtn_AddNewContract_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick ActBtn_AddNewContract

End Sub

'Mouse Controls Starts

Private Sub ActBtn_addnewcontract_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Me.ActBtn_AddNewContract
    'Me.ActBtn_AddNewContract.BorderStyle = fmBorderStyleSingle
    
    highlight Me.ActBtn_Save

End Sub



Private Sub ActBtn_AddNewContract_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight ActBtn_AddNewContract

End Sub

Private Sub ActBtn_DeleteContract_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick ActBtn_DeleteContract

End Sub

Private Sub ActBtn_DeleteContract_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Me.ActBtn_DeleteContract

End Sub


Private Sub ActBtn_DeleteContract_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight ActBtn_DeleteContract
    
End Sub

Private Sub ActBtn_NewContract_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick ActBtn_NewContract

End Sub

Private Sub ActBtn_NewContract_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Me.ActBtn_NewContract
    Me.ActBtn_NewContract.BorderStyle = fmBorderStyleSingle
    
    
End Sub

Private Sub ActBtn_NewContract_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight ActBtn_NewContract

End Sub

Private Sub ActBtn_Save_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick ActBtn_Save

End Sub

Private Sub ActBtn_Save_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight ActBtn_Save

End Sub

Private Sub Btn_Amendments_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_Amendments

End Sub

Private Sub Btn_Amendments_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Me.MultiPage1.Value = 2 Then
        highlight Btn_Amendments
    Else
        Exit_highlight Btn_Amendments
    End If

End Sub

Private Sub Btn_AssignPco1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_AssignPco1

End Sub

Private Sub Btn_AssignPco1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    highlight Me.Btn_AssignPco1

End Sub

Private Sub Btn_AssignPco1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_AssignPco1

End Sub


Private Sub Btn_AssignPco2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_AssignPco2

End Sub

Private Sub Btn_AssignPco2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_AssignPco2

End Sub

Private Sub Btn_AssignPco2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  highlight Me.Btn_AssignPco2
  
End Sub


Private Sub Btn_ChangeDirectory_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_ChangeDirectory

End Sub

Private Sub Btn_ChangeDirectory_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    highlight Me.Btn_ChangeDirectory

End Sub

Private Sub Btn_ChangeDirectory_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Btn_ChangeDirectory


End Sub

Private Sub Btn_CLMSlink_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Onclick Btn_CLMSlink
    
End Sub

Private Sub Btn_CLMSlink_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_CLMSlink

End Sub

Private Sub Btn_DeviationDetails_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_DeviationDetails

End Sub

Private Sub Btn_DeviationDetails_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Me.MultiPage1.Value = 1 Then
        highlight Btn_DeviationDetails
    Else
        Exit_highlight Btn_DeviationDetails
    End If

End Sub

Private Sub Btn_GeneralInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_GeneralInfo

End Sub



Private Sub Btn_GeneralInfo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Me.MultiPage1.Value = 0 Then
        highlight Btn_GeneralInfo
    Else
        Exit_highlight Btn_GeneralInfo
    End If

End Sub

Private Sub Btn_Link_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_Link
End Sub

Private Sub Btn_Link_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.Btn_Link
    
End Sub


Private Sub Btn_CLMSlink_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    highlight Me.Btn_CLMSlink

End Sub


Private Sub Btn_Link_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_Link

End Sub

Private Sub Btn_Next_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_Next

End Sub

Private Sub Btn_Next_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Exit_highlight Btn_Previous
    highlight Btn_Next

End Sub

Private Sub Btn_Next_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Onclick Btn_Next

End Sub

Private Sub Btn_OpenFile_Click()

    For i = 0 To Me.ListBox2.ListCount - 1
    
        If Me.ListBox2.Selected(i) = True And Me.ListBox2.List(i, 0) <> "" Then
        
            HLink = Me.ListBox2.List(i, 1) & Me.ListBox2.List(i, 0)
            CreateObject("WScript.Shell").Run Chr(34) & HLink & Chr(34)
            Exit For
        End If
    
    Next i

End Sub


Private Sub Btn_OpenFile_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_OpenFile

End Sub


Private Sub Btn_OpenFile_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_OpenFile

End Sub

Private Sub Btn_OpenFolder_Click()

    Dim Obj As Object
    
    Set Obj = CreateObject("Scripting.FileSystemObject")
    
    If Me.Field_33.Value <> "" And Obj.FolderExists(Me.Field_33.Value & "\") Then
    
        CreateObject("WScript.Shell").Run Chr(34) & Me.Field_33.Value & "\" & Chr(34)
    End If
    
End Sub

Private Sub Btn_OpenFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.Btn_OpenFile

End Sub



Private Sub Btn_OpenFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Onclick Btn_OpenFolder

End Sub


Private Sub Btn_OpenFolder_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Exit_highlight Btn_OpenFolder

End Sub


Private Sub Btn_OpenFolder_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.Btn_OpenFolder
    
End Sub



Private Sub Btn_Previous_Click()
Dim settings As New ExclClsSettings
 
settings.TurnOn
ThisWorkbook.Activate
    If ActiveCell.row = 18 Then
        ActiveSheet.Cells(rows.count, 4).End(xlUp).Select
    Else
        ActiveCell.Offset(-1).Select
    End If
    openContractEdit
    
End Sub

Private Sub Btn_Next_Click()
Dim settings As New ExclClsSettings

settings.TurnOn

ThisWorkbook.Activate
    If ActiveCell.row = ActiveSheet.Cells(rows.count, 4).End(xlUp).row Then
        ActiveSheet.Cells(18, 4).Select
    Else
        ActiveCell.Offset(1).Select
    End If
    openContractEdit
    
End Sub

Private Sub Btn_Previous_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_Previous

End Sub

Private Sub Btn_Previous_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        
    Exit_highlight Btn_Next
    highlight Btn_Previous

End Sub

Private Sub Btn_Previous_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

     Onclick Btn_Previous

End Sub

Private Sub Btn_RefTermDates_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_RefTermDates

End Sub



Private Sub Btn_RefTermDates_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    highlight Me.Btn_RefTermDates

End Sub


Private Sub ActBtn_Save_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Exit_highlight Me.ActBtn_Save
    Me.ActBtn_Save.BorderStyle = fmBorderStyleSingle
    'Me.ActBtn_Save.BackStyle = fmBackStyleTransparent
    
    highlight Me.ActBtn_AddNewContract
    
End Sub



Private Sub Btn_RefTermDates_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_RefTermDates

End Sub

Private Sub Btn_RenewalInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_RenewalInfo

End Sub

Private Sub Btn_RenewalInfo_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Me.MultiPage1.Value = 3 Then
        highlight Btn_RenewalInfo
    Else
        Exit_highlight Btn_RenewalInfo
    End If

End Sub

Private Sub Btn_Status_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_Status

End Sub

Private Sub Btn_Status_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If Me.MultiPage1.Value = 4 Then
        highlight Btn_Status
    Else
        Exit_highlight Btn_Status
    End If

End Sub

Private Sub Btn_UnassignPco1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_UnassignPco1

End Sub

Private Sub Btn_UnassignPco2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Onclick Btn_UnassignPco2

End Sub

Private Sub Btn_UnassignPco1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_UnassignPco1

End Sub

Private Sub Btn_UnassignPco2_Mouseup(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Exit_highlight Btn_UnassignPco2

End Sub

Private Sub Btn_UnassignPco1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    highlight Me.Btn_UnassignPco1

End Sub



Private Sub Btn_UnassignPco2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     
     highlight Me.Btn_UnassignPco2

End Sub



Private Sub ComboBox1_Change()

    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
  
    If ComboBox2.Value = "" Then ComboBox2.Value = 0
    
    If ComboBox1.Value = "" Then ComboBox1.Value = 0
    
    If (ComboBox1.Value <> 0 Or ComboBox2.Value <> 0) And (Field_25.Value = 0 Or Field_25.Value = "") Then Field_25.Value = 1
    
    If (ComboBox1.Value = 0 And ComboBox2.Value = 0) And (Field_25.Value <> 0) Then Field_25.Value = 0
    
    Field_26.Caption = Round(CDbl(ComboBox1.Value) + CDbl(ComboBox2.Value) / 12, 2)
    
    Btn_RefTermDates.Visible = True
    
    settings.Restore

End Sub

Private Sub ComboBox2_Change()
    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    If ComboBox2.Value = "" Then ComboBox2.Value = 0
    
    If ComboBox1.Value = "" Then ComboBox1.Value = 0
    
    If (ComboBox1.Value <> 0 Or ComboBox2.Value <> 0) And (Field_25.Value = 0 Or Field_25.Value = "") Then Field_25.Value = 1
    
    If (ComboBox1.Value = 0 And ComboBox2.Value = 0) And (Field_25.Value <> 0) Then Field_25.Value = 0
      
    Field_26.Caption = Round(CDbl(ComboBox1.Value) + CDbl(ComboBox2.Value) / 12, 2)
    
    Btn_RefTermDates.Visible = True
    
    settings.Restore

End Sub

Private Sub ComboBox3_Change()
    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
      
    If ComboBox4.Value = "" Then ComboBox2.Value = 0
        
    If ComboBox3.Value = "" Then ComboBox3.Value = 0
      
    Field_30.Caption = Round(CDbl(ComboBox3.Value) + CDbl(ComboBox4.Value) / 12, 2)
     
    Btn_RefTermDates.Visible = True
    
    settings.Restore
End Sub

Private Sub ComboBox4_Change()

    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff

    If ComboBox4.Value = "" Then ComboBox4.Value = 0
    
    If ComboBox3.Value = "" Then ComboBox3.Value = 0
    
    Field_30.Caption = Round(CDbl(ComboBox3.Value) + CDbl(ComboBox4.Value) / 12, 2)
    
    Btn_RefTermDates.Visible = True

    settings.Restore
    
End Sub

Private Sub Field_34_Change()
    
    If Me.Field_34.Value <> "" Then Me.Priority.Caption = "  " & Me.Field_34.Value & " Priority Contract"

End Sub

Private Sub Field_8_Change()

    If Field_8.Value = "" Then
       
       Field_8.BorderStyle = fmBorderStyleSingle
        
        Field_8.BorderColor = ThemeUf.EmptyTB.BorderColor
    
    Else
        
        Field_8.BorderColor = ThemeUf.FilledTB.BorderColor
    
    End If

End Sub



Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    
    
    Dim curFile As String
    Dim r As Long
    Debug.Print Data.Files.count
    
    
    
    
    For i = 1 To Data.Files.count
    
        curFile = Data.Files(i)
        r = Me.ListBox2.ListCount
        
        Me.ListBox2.AddItem
        Me.ListBox2.List(r, 0) = getFName(curFile)
        Me.ListBox2.List(r, 1) = Replace((curFile), getFName(curFile), "")
        
    Next i
    
    'If Me.Field_33.Value = "" Then SelectFolderLoc
    
End Sub

Public Sub SelectFolderLoc()
'https://dhhsemployees/sites/Operations/Procurement/Contract%20Administration/_layouts/15/guestaccess.aspx?guestaccesstoken=Yje9x0CZbnzdIcwWABmCvnYoCUo4KG%2f3C5hMezyw0Gc%3d&docid=2_0f82b9ae0064a4e1ea71e333200323a08&rev=1
Me.Field_33.Enabled = True
Dim dialogBox As FileDialog

Dim Obj As Object

Set Obj = CreateObject("Scripting.fileSystemObject")

Dim initialValue As String

    initialValue = Me.Field_33.Value
    
    Set dialogBox = Application.FileDialog(msoFileDialogFolderPicker)

'Do not allow multiple files to be selected
    dialogBox.AllowMultiSelect = False

'Set the title of the DialogBox
    dialogBox.Title = "Select Folder to Save This Files"

'Set the default folder to open
    If initialValue = "" Or Not Obj.FolderExists(initialValue & "\") Then
        
        If Not Obj.FolderExists(initialValue & "\") Then MsgBox ("This Folder Location Does Not Exist")
        
        dialogBox.InitialFileName = "C:\Users"
    Else
        dialogBox.InitialFileName = Me.Field_33.Value & "\"
    End If

'Apply file filters - use ; to separate filters for the same name

'Show the dialog box and output full file name
    If dialogBox.Show = -1 Then
    
        Me.Field_33.Value = dialogBox.SelectedItems(1)
    End If
      
    Dim strfile As String, row As Integer, yesNo As String
            
            
            If Me.Field_33.Value <> "" Then
            
                
            End If
            
            If Me.Field_33.Value <> initialValue Then
                
                yesNo = MsgBox("Do you Want to move the contents Of previous folder into the New Destination?", vbYesNo, "Move Files?")
            
                If yesNo = vbNo Then Me.ListBox2.Clear
                
                row = Me.ListBox2.ListCount
                
                Me.Btn_ChangeDirectory.Visible = True
                
                strfile = Dir(Me.Field_33.Value & "\")
                
                Do While Len(strfile) > 0
                    'Debug.Print .getFName(strfile), strfile
                    Me.ListBox2.AddItem
                    Me.ListBox2.List(row, 0) = strfile
                    Me.ListBox2.List(row, 1) = Me.Field_33.Value & "\"
                    row = Me.ListBox2.ListCount
                    'row = row + 1
                    strfile = Dir
                Loop
            
            End If
    
    Me.Btn_ChangeDirectory.Visible = True
    
    Me.Field_33.Enabled = False
    
End Sub

Public Function getFName(curFile As String) As String

    Dim strLen As Long
    Dim i As Long
    Dim fileNameLen As Long
    
    fileNameLen = 0
    
    strLen = Len(curFile)
    
    For i = strLen To 1 Step -1
        
        If Mid(curFile, i, 1) = "\" Then
        
            getFName = Right(curFile, fileNameLen)
            
            Exit For
        End If
        fileNameLen = fileNameLen + 1

    Next i

End Function
Private Sub Btn_ChangeDirectory_Click()
    
    SelectFolderLoc

End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.ActBtn_Save
    highlight Me.ActBtn_NewContract
    highlight Me.ActBtn_AddNewContract

End Sub

Private Sub PCO_Change()
    
    If Me.PCO.Value = "" Then
        
        PCO.BorderStyle = fmBorderStyleSingle
            
        PCO.BorderColor = EmptyTB.BorderColor
        
    Else
    
        PCO.BorderColor = ThemeUf.FilledTB.BorderColor
        
    End If
    
    
End Sub

Private Sub pg1_title_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.ActBtn_Save
    highlight Me.ActBtn_NewContract
    Exit_highlight Me.Btn_Next

End Sub


Private Sub pg2_title_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.ActBtn_Save
    highlight Me.ActBtn_NewContract

End Sub


Private Sub pg3_title_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.ActBtn_Save
    highlight Me.ActBtn_NewContract
        Exit_highlight Me.Btn_Next

End Sub


Private Sub pg4_title_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.ActBtn_Save
    highlight Me.ActBtn_NewContract
    Exit_highlight Me.Btn_Next

End Sub


Private Sub pg5_title_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.ActBtn_Save
    highlight Me.ActBtn_NewContract
    Exit_highlight Me.Btn_Next
    
End Sub


Private Sub Bg_pg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    checkHighlight
    
    highlight Me.ActBtn_DeleteContract
    Exit_highlight Btn_Next
    Exit_highlight Btn_Previous
    
    highlight Me.ActBtn_AddNewContract
    
    Exit_highlight Me.Btn_Link
    Me.Btn_Link.BorderStyle = fmBorderStyleSingle
    'Me.Btn_Link.BackStyle = fmBackStyleTransparent

    
    Exit_highlight Me.Btn_AssignPco1
    Me.Btn_AssignPco1.BorderStyle = fmBorderStyleSingle
    'Me.Btn_AssignPco1.BackStyle = fmBackStyleTransparent
    
    Exit_highlight Me.Btn_AssignPco2
    Me.Btn_AssignPco2.BorderStyle = fmBorderStyleSingle
    'Me.Btn_AssignPco2.BackStyle = fmBackStyleTransparent
    
    Exit_highlight Me.Btn_UnassignPco1
    Me.Btn_UnassignPco1.BorderStyle = fmBorderStyleSingle
    'Me.Btn_UnassignPco1.BackStyle = fmBackStyleTransparent
    
    Exit_highlight Me.Btn_UnassignPco2
    Me.Btn_UnassignPco2.BorderStyle = fmBorderStyleSingle
    'Me.Btn_UnassignPco2.BackStyle = fmBackStyleTransparent
    
    highlight Me.ActBtn_NewContract
    Exit_highlight Me.Btn_Next

highlight Me.ActBtn_Save
    
End Sub

Private Sub Bg_pg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
checkHighlight
    Exit_highlight Btn_Next
    Exit_highlight Btn_Previous
highlight Me.ActBtn_Save
highlight Me.ActBtn_NewContract
highlight Me.ActBtn_AddNewContract
End Sub

Private Sub Bg_pg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
checkHighlight
    Exit_highlight Btn_Next
    Exit_highlight Btn_Previous
highlight Me.ActBtn_Save
highlight Me.ActBtn_AddNewContract
highlight Me.ActBtn_NewContract
End Sub



Private Sub Bg_pg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
        checkHighlight
    Exit_highlight Me.Btn_RefTermDates
    Me.Btn_RefTermDates.BorderStyle = fmBorderStyleSingle
    'Me.Btn_RefTermDates.BackStyle = fmBackStyleTransparent
    Exit_highlight Btn_Next
    highlight Me.ActBtn_AddNewContract
    Exit_highlight Btn_Previous
highlight Me.ActBtn_Save
highlight Me.ActBtn_NewContract

End Sub


Private Sub Bg_pg5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
checkHighlight
    Exit_highlight Btn_Next
    Exit_highlight Btn_Previous
    highlight Me.ActBtn_Save
    Exit_highlight Me.Btn_CLMSlink
    Exit_highlight Me.Btn_ChangeDirectory
    Exit_highlight Me.Btn_OpenFile
    Exit_highlight Me.Btn_OpenFolder
    highlight Me.ActBtn_NewContract
    highlight Me.ActBtn_AddNewContract
    Me.Btn_CLMSlink.BorderStyle = fmBorderStyleSingle
    'Me.Btn_CLMSlink.BackStyle = fmBackStyleTransparent

End Sub


Private Sub Priority_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    highlight Me.ActBtn_Save
    highlight Me.ActBtn_NewContract
    Exit_highlight Me.Btn_Next
    Exit_highlight Me.Btn_Previous

End Sub


Private Sub Btn_GeneralInfo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

If Me.MultiPage1.Value <> 0 Then
    highlight Me.Btn_GeneralInfo
End If

If Me.MultiPage1.Value <> 1 Then
    Exit_highlight Me.Btn_DeviationDetails
End If

End Sub



Private Sub Btn_DeviationDetails_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Me.MultiPage1.Value <> 0 Then
        Exit_highlight Me.Btn_GeneralInfo
    End If
    If Me.MultiPage1.Value <> 1 Then
        highlight Me.Btn_DeviationDetails
    End If
    If Me.MultiPage1.Value <> 2 Then
    Exit_highlight Me.Btn_Amendments
    End If
    
End Sub


Private Sub Btn_Amendments_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Me.MultiPage1.Value <> 1 Then
        Exit_highlight Me.Btn_DeviationDetails
    End If
    If Me.MultiPage1.Value <> 2 Then
        highlight Me.Btn_Amendments
    End If
    If Me.MultiPage1.Value <> 3 Then
        Exit_highlight Me.Btn_RenewalInfo
    End If
    
End Sub

Private Sub Btn_RenewalInfo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If Me.MultiPage1.Value <> 2 Then
    Exit_highlight Me.Btn_Amendments
    End If
    If Me.MultiPage1.Value <> 3 Then
        highlight Me.Btn_RenewalInfo
    End If
    If Me.MultiPage1.Value <> 4 Then
        Exit_highlight Me.Btn_Status
    End If
End Sub

Private Sub Btn_Status_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me.MultiPage1.Value <> 3 Then
        Exit_highlight Me.Btn_RenewalInfo
    End If
    If Me.MultiPage1.Value <> 4 Then
        highlight Me.Btn_Status
    End If
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Exit_highlight Btn_Previous
Exit_highlight Btn_Next
checkHighlight

highlight Me.ActBtn_AddNewContract
highlight Me.ActBtn_Save
highlight Me.ActBtn_NewContract

End Sub


'Mouse Controls ends


Private Sub Btn_RefTermDates_Click()


    Dim ted             As Date 'term End Date
    Dim Nor             As Integer ' No. Of Renewals
    Dim Erd             As Double ' each renewal duration
    Dim Tsd             As Date 'Term start date
    Dim Ext             As Double
    
    
        
    If Field_23.Value = "" Or Field_24.Value = "" Then
        
        If Field_23.Value = "" Then
            
            Field_23.BorderColor = &HFF&
            
            Field_23.BorderStyle = fmBorderStyleSingle
        
        End If
        
        If Field_24.Value = "" Then
            
            Field_24.BorderColor = &HFF&
            
            Field_24.BorderStyle = fmBorderStyleSingle
        
        End If
        
        Exit Sub
    
    End If
        
        Tsd = Field_23.Value
        
        ted = Field_24.Value
        
        Nor = IIf(Field_25.Value = "", 0, Field_25.Value)
        
        If Field_26.Caption = "" Then
            
            Erd = 0
        
        Else
            
            Erd = CDbl(Field_26.Caption)
        
        End If
        
        If Field_30.Caption = "" Then
            
            Ext = 0
        
        Else
            
            Ext = CDbl(Field_30.Caption)
        
        End If
        
    CalcTerm Tsd, ted, Nor, Erd, Ext
    
    ThisWorkbook.Activate
      
      With ListBox1
            
            .ColumnCount = 4
            
            .ColumnWidths = "70;70;70;70"
            
            .RowSource = "Sheet1!" & Sheet17.Range("AB3").CurrentRegion.Offset(1).Address
        
        End With
    
    Btn_RefTermDates.Visible = False



End Sub


Private Sub ActBtn_NewContract_Click()
        
    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    Dim wb As Workbook
    Dim staffSheet As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim fileLocation As String
    Dim pk As String
    Dim columnNum As Long
    Dim cntrls As Control
    Dim duration() As String
    Set staffSheet = Sheet6
    Dim pos As String
    Dim maxnum As String
    Dim BeforeUpdate As String
    Dim pkNo As Long
    
    If PCO.Value = "" Or Field_8.Value = "" Or Field_3.Value = "" Then
        
        If PCO.Value = "" Then
            
            EmptyField PCO
        
        Else
            
            baseField PCO
        
        End If
        
        If Field_3.Value = "" Then
           
            EmptyField Field_3
        
        Else
            
            baseField Field_3
        
        End If
        
        If Field_8.Value = "" Then
           
            EmptyField Field_8
        
        Else
            
            baseField Field_8
        
        End If
        
        Exit Sub
    
    End If
    
        For i = 2 To staffSheet.Cells(1, 4).End(xlDown).row
        
        If PCO.Value = staffSheet.Cells(i, 4).Value Then
                                
                pos = staffSheet.Cells(i, 6).Value
            
            Exit For
        
        End If
    
    Next i
    
    On Error GoTo Handler
       
    Set ws = Sheet8
    
    getUpdatedData
    
    'Assign different PK if temppco is assigned
    
    pkNo = Application.WorksheetFunction.Max(ws.Range("AR2:AR" & ws.Range("AR2").End(xlDown).row)) + 1
    
        
    pk = pos & "_" & pkNo
    
    If ws.Cells(1, 1).Value = "" Then
        
        i = ws.Cells(1, 1).End(xlDown).row
    
    Else
        
        i = ws.Cells(1, 1).End(xlDown).row + 1
    
    End If
    
    For Each cntrls In Me.Controls
        
        On Error Resume Next
        
        If Len(cntrls.name) > 7 Then
            
            columnNum = CInt(Right(cntrls.name, 2))
        
        Else
            
            columnNum = CInt(Right(cntrls.name, 1))
        
        End If
        
        On Error GoTo 0
        
        If columnNum = 1 Or columnNum = 2 Then
            
            If columnNum <> 1 Then BeforeUpdate = BeforeUpdate & ws.Cells(i, columnNum)
            
            ws.Cells(i, 1).Value = pk
            
            ws.Cells(i, 2).Value = PCO.Value
        
        ElseIf TypeName(cntrls) = "TextBox" Or TypeName(cntrls) = "ComboBox" And Left(cntrls.name, 5) = "Field" Then
            
            ws.Cells(i, columnNum).Value = cntrls.Value
            
            If columnNum = 26 And Sheet8.Cells(i, columnNum).Value < 1 Then
                
                ReDim duration(1 To 2)
                
                duration = Split(cntrls.Value, " ")
                
                If duration(1) = "month" Then
                    
                    Sheet8.Cells(i, columnNum).Value = CInt(duration(0)) * 12
                
                Else
                    
                    Sheet8.Cells(i, columnNum).Value = CInt(duration(0))
                
                End If
            
            Else
                
                Sheet8.Cells(i, columnNum).Value = cntrls.Value
            
            End If
            
            If Security.Caption <> "Admin" Then cntrls.Enabled = False
        
        ElseIf TypeName(cntrls) = "Label" And Left(cntrls.name, 5) = "Field" And InStr(1, cntrls.name, "Field_Header", vbTextCompare) = 0 Then
              ws.Cells(i, columnNum).Value = cntrls.Caption
        
        End If
    
    Next cntrls
    
    CalcDates ws, i
    
    typ = Me.Field_3.Value
    
    ws.Cells(i, "AS").Value = "No"
    ws.Cells(i, "AT").Value = ""
    
    syncData
    
    updateLog ThisWorkbook, "Added New Contract Record:" & pk, "Add New Contracts"

    Select Case (typ)
        
        Case "Current_Active"
            Sheet2.Range("A3").Value = True
            
        Case "Current_Renewal"
            Sheet23.Range("A3").Value = True
            
        Case "Current_Amendment"
            Sheet4.Range("A3").Value = True
            
        Case "Current_Extension"
            Sheet24.Range("A3").Value = True
            
        Case "New Procurement_(No Existing)"
            Sheet3.Range("A3").Value = True
            
        Case "New Procurement_(Replace Existing)"
            Sheet5.Range("A3").Value = True
            
        Case "Closed"
            Sheet20.Range("A3").Value = True
            
        Case "Deviation"
            Sheet10.Range("A3").Value = True
            
    End Select
        
    settings.Restore
    
    Exit Sub
    
Handler:
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
  
        
        updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "Add New Contracts : Unsucessful"
    
    settings.Restore
End Sub




Private Sub ActBtn_DeleteContract_Click()

 Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    Dim wb As Workbook
    Dim staffSheet As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim fileLocation As String
    Dim pk As String
    Dim columnNum As Long
    Dim cntrls As Control
    Dim duration() As String
    Set staffSheet = Sheet6
    Dim pos As String
    Dim maxnum As String
    Dim BeforeUpdate As String
    Dim reason As String
    Dim typ As String
    Dim yesNo As String
    
    If Sheet12.Range("pName").Value <> Field_2.Caption And Sheet12.Range("Security").Value <> "Admin" Then
    
        MsgBox "Cannot Make changes to other PCO profiles!"
        
        Unload Me
        
        settings.Restore
        
        Exit Sub
        
    End If
    
    yesNo = MsgBox(" Do you Want to Delete this Contract?", vbYesNo, "Confirm Delete?")
    
    If yesNo = vbNo Then Exit Sub
    
    reason = InputBox("Why are you deleting this contract?", "Reason For Deleting Contract Record")
    
        
    On Error GoTo Handler
    
    
    Set ws = Sheet8
    
    If ws.Cells(1, 1).Value = "" Then
        
        i = ws.Cells(1, 1).End(xlDown).row
    
    Else
        
        i = ws.Cells(1, 1).End(xlDown).row + 1
    
    End If
    
    typ = Me.Field_3.Value
    
    For i = 1 To ws.Cells(1, 1).End(xlDown).row
        
        If ws.Cells(i, db_primaryKey).Value = Primary_Key.Caption Then
            
            ws.Cells(i, "AS").Value = "Yes"
            
            ws.Cells(i, "AT").Value = ""
            
            Exit For
        
        End If
    
    Next i
    
    updateLog ThisWorkbook, "Deleted Record:" & Primary_Key.Caption, "Deleted Contract" & reason
    
    syncData
   
    Select Case (typ)
        
        Case "Current_Active"
            Sheet2.Range("A3").Value = True
            
        Case "Current_Renewal"
            Sheet23.Range("A3").Value = True
            
        Case "Current_Amendment"
            Sheet4.Range("A3").Value = True
            
        Case "Current_Extension"
            Sheet24.Range("A3").Value = True
            
        Case "New Procurement_(No Existing)"
            Sheet3.Range("A3").Value = True
            
        Case "New Procurement_(Replace Existing)"
            Sheet5.Range("A3").Value = True
            
        Case "Closed"
            Sheet20.Range("A3").Value = True
            
        Case "Deviation"
            Sheet10.Range("A3").Value = True
            
    End Select
    
    Unload Me
    
    settings.Restore
    
    Exit Sub
    
Handler:

        
    updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "Delete Contracts : Unsucessful"
    
    settings.Restore

End Sub



Private Sub Btn_AssignPco1_Click()

    If Field_11.Value = Field_2.Caption Then
        
        MsgBox "Temp PCO Cannot be same as the PCO"
        
        Exit Sub
    
    End If
    
    With Me
    
        '.TempPco1.Caption = "True"
        
        '.PCO.Value = .Field_11.Value
        
        'ActBtn_NewContract_Click
        
        '.ActBtn_NewContract.Visible = False
        
        .Field_11.Enabled = False
        
        .Btn_UnassignPco1.Visible = True
        
        .Btn_AssignPco1.Visible = False
        
        .PCO.Value = ""
        
        '.TempPco1.Caption = "False"
    
    End With
    
End Sub



Private Sub Btn_AssignPco2_Click()

    If Field_11.Value = Field_2.Caption Then
        
        MsgBox "Temp PCO Cannot be same as the PCO"
        
        Exit Sub
    
    End If
    
    With Me
    
        .TempPco2.Caption = "True"
        
        '.PCO.Value = .Field_12.Value
        
        'ActBtn_NewContract_Click
        
        '.ActBtn_NewContract.Visible = False
        .Field_12.Enabled = False
        
        .Btn_UnassignPco2.Visible = True
        
        .Btn_AssignPco2.Visible = False
        
        '.PCO.Value = ""
        
        .TempPco2.Caption = "False"
    
    End With
    
End Sub


Private Sub Btn_UnassignPco1_Click()

Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    Dim wb As Workbook
    Dim staffSheet As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim fileLocation As String
    Dim pk As String
    Dim columnNum As Long
    Dim cntrls As Control
    Dim duration() As String
    Set staffSheet = Sheet6
    Dim remarkPK() As String
    
        
'    Set ws = Sheet8
'
'    pk = Primary_Key.Caption
'
'    For i = 2 To ws.Cells(1, 1).End(xlDown).row
'
'    ReDim remarkPK(1 To 2)
'
'    remarkPK = Split(ws.Cells(i, 28).Value, ";")
'
'        If ws.Cells(i, 28).Value = "" Then
'
'        ElseIf remarkPK(0) = pk And ws.Cells(i, 2).Value = Field_11.Value Then
'
'            ws.Cells(i, "AS").Value = "Yes"
'
'            ws.Cells(i, "AT").Value = ""
'
'            syncData
'
'            Exit For
'
'        End If
'
'    Next i
        
    TempPco1.Caption = "True"
    
    Field_11.Value = ""
    
    Btn_UnassignPco1.Visible = False
    
    ActBtn_Save_Click
    
    Field_11.Enabled = True
    
    TempPco1.Caption = "False"
    
    settings.Restore
    
End Sub




Private Sub Btn_UnassignPco2_Click()

    Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    Dim wb As Workbook
    Dim staffSheet As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    Dim fileLocation As String
    Dim pk As String
    Dim columnNum As Long
    Dim cntrls As Control
    Dim duration() As String
    Set staffSheet = Sheet6
    Dim remarkPK() As String

    
'    Set ws = Sheet8
'
'    pk = Primary_Key.Caption
'
'    For i = 2 To ws.Cells(1, 1).End(xlDown).row
'
'    ReDim remarkPK(1 To 2)
'
'    remarkPK = Split(ws.Cells(i, 28).Value, ";")
'
'        If ws.Cells(i, 28).Value = "" Then
'
'        ElseIf remarkPK(0) = pk And ws.Cells(i, 2).Value = Field_12.Value Then
'
'            ws.Cells(i, "AS").Value = "Yes"
'
'            ws.Cells(i, "AT").Value = ""
'
'            Exit For
'
'        End If
'
'    Next i
    
    TempPco2.Caption = "True"
    
    Field_12.Value = ""
    
    Btn_UnassignPco2.Visible = False
    
    ActBtn_Save_Click
    
    Field_12.Enabled = True
    
    TempPco2.Caption = "False"
    
    settings.Restore

End Sub


Private Sub ActBtn_Save_Click()
Dim settings As New ExclClsSettings
    'Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    
    settings.TurnOff
    
    Dim staffSheet As Worksheet
    Dim ws As Worksheet, localWs As Worksheet
    Dim i As Long, j As Long
    Dim fileLocation As String
    Dim pk As String
    Dim columnNum As Long
    Dim cntrls As Control
    Dim duration() As String
    Set staffSheet = Sheet6
    Dim rng As Range
    Dim beforeUpdateVals As String
    Dim tempPk As String
    Dim yesNo As String
    
    Dim tempRow As Long
    
    Set localWs = Sheet8
    
    If Sheet12.Range("pName").Value <> Field_2.Caption And Sheet12.Range("Security").Value <> "Admin" And Sheet12.Range("pName").Value <> Field_11.Value And Sheet12.Range("pName").Value <> Field_12.Value Then
    
        MsgBox "Cannot Make changes to other PCO profiles!"
        
        Unload Me
        
        settings.Restore
        
        Exit Sub
        
    End If
    
    Dim EmptyCnt As Integer
    
    EmptyCnt = 0
    
    If Field_13.Value = "Contract Executed (Final Status)" And Field_13.Enabled = False Then
        
        For Each cntrls In Controls
            
            For Each rng In Sheet17.Range("MandFields")
                
                If cntrls.name = rng.Value And cntrls.Value = "" Then
                   
                   cntrls.BorderStyle = fmBorderStyleSingle
                   
                   cntrls.BorderColor = &HFF&
                   
                   EmptyCnt = EmptyCnt + 1
                
                Else
                    
                    cntrls.BorderStyle = fmBorderStyleNone
                
                End If
            
            Next rng
        
        Next cntrls
        
        If EmptyCnt > 0 Then
            
            MsgBox ("Please fill out all the mandatory fields and Try again!")
            
            Exit Sub
        
        End If
    
    End If
    
    
    On Error GoTo Handler:

    Set ws = Sheet8
    
    pk = Primary_Key.Caption
    
    tempPk = ""
    
    beforeUpdateVals = "Primary_Key:" & pk
    
'    If InStr(1, pk, "Temp", vbTextCompare) > 0 Then
'
'    tempPk = Me.TempContrNum.Caption
'
'    For Each rng In Sheet8.Range("A1:A" & Sheet8.Range("A1").End(xlDown).row)
'
'        If rng.Value = tempPk Then
'
'            tempRow = rng.row
'
'            Exit For
'
'        End If
'
'    Next rng
'
'    End If
    
    i = CInt(Me.db_Row.Caption)
    
        If ws.Cells(i, 1).Value = pk Then
        
                        
            For Each cntrls In Me.Controls
                
                On Error Resume Next
                
                If Len(cntrls.name) > 7 Then
                    
                    columnNum = CInt(Right(cntrls.name, 2))
                
                Else
                    
                    columnNum = CInt(Right(cntrls.name, 1))
                
                End If
                
                If columnNum <= 2 Then GoTo NextControl
                On Error GoTo 0
                
                If TypeName(cntrls) = "TextBox" Or TypeName(cntrls) = "ComboBox" And Left(cntrls.name, 5) = "Field" Then
                      
                 'Debug.Print cntrls.name, cntrls.Value
                      
                      beforeUpdateVals = beforeUpdateVals & "|" & ws.Cells(1, columnNum).Value & ";" & ws.Cells(i, columnNum).Value
                      
                      ws.Cells(i, columnNum).Value = cntrls.Value
                      
                    'If tempPk <> "" Then ws.Cells(tempRow, columnNum).Value = cntrls.Value
                    
                    If columnNum = 26 And Sheet8.Cells(i, columnNum).Value < 1 Then
                        
                        ReDim duration(1 To 2)
                        
                        duration = Split(cntrls.Value, " ")
                        
                        If duration(1) = "month" Then
                            
                            ws.Cells(i, columnNum).Value = CInt(duration(0)) * 12
                            
                           ' If tempPk <> "" Then ws.Cells(tempRow, columnNum).Value = CInt(duration(0)) * 12
                        
                        Else
                            
                            ws.Cells(i, columnNum).Value = CInt(duration(0))
                            
                           'If tempPk <> "" Then ws.Cells(tempRow, columnNum).Value = CInt(duration(0))
                        
                        End If
                    
                    ElseIf columnNum = 33 Then
                    
                       
                        ws.Cells(i, columnNum).Value = ""
                        
                    Else
                        
                        ws.Cells(i, columnNum).Value = cntrls.Value
                        
                       'If tempPk <> "" Then ws.Cells(tempRow, columnNum).Value = cntrls.Value
                    
                    End If
                    
                    If Security.Caption <> "Admin" Then cntrls.Enabled = False
                
                ElseIf TypeName(cntrls) = "Label" And Left(cntrls.name, 5) = "Field" And InStr(1, cntrls.name, "Field_Header", vbTextCompare) = 0 Then
                      
                      'Debug.Print cntrls.name, cntrls.Caption
                      
                      beforeUpdateVals = beforeUpdateVals & "|" & ws.Cells(1, columnNum).Value & ";" & ws.Cells(i, columnNum).Value
                      
                      ws.Cells(i, columnNum).Value = cntrls.Caption
                    
'                      If columnNum <> 1 And columnNum <> 2 Then
'
'                       'If tempPk <> "" Then ws.Cells(tempRow, columnNum).Value = cntrls.Caption
'
'                      End If
                    
                End If
NextControl:
            Next cntrls
            
        ws.Cells(i, 46).Value = ""
        
        CalcDates ws, i
        
        Dim Obj As Object, rootFolder As String, contractFolder
        
        rootFolder = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "PCO Contract Files")
        
        Set Obj = CreateObject("Scripting.fileSystemObject")
        
        If Me.ListBox2.ListCount > 0 Then
        
            For i = 0 To Me.ListBox2.ListCount - 1
                
                If Not Obj.FolderExists(rootFolder) Then
                
                    Obj.CreateFolder (rootFolder)
                
                End If
                
                contractFolder = rootFolder & "\" & Me.Primary_Key.Caption & " " & Me.Field_8.Value
                
                strfile = Dir(rootFolder & "\")
                
                Do While Len(strfile) > 0
                    
                    If InStr(1, strfile, Me.Primary_Key.Caption, vbTextCompare) > 0 Then
                        
                       Name strfile As contractFolder
                        
                    End If
                    
                    strfile = Dir
                Loop
                
                
                
                If Me.ListBox2.List(i, 1) <> contractFolder & "\" Then
                
                    If Not Obj.FolderExists(contractFolder) Then
                        
                        Obj.CreateFolder (contractFolder)
                    
                    End If
                    
                    FileCopy Me.ListBox2.List(i, 1) & Me.ListBox2.List(i, 0), contractFolder & "\" & Me.ListBox2.List(i, 0)
                
                End If
                
            Next i
        
        End If
        
        
        
        End If
        
    syncData
    
    updateLog ThisWorkbook, "Edited Contract Rec:" & beforeUpdateVals, "Save Updates Procedure"
    
    MsgBox ("Changes Saved Successfully!")
    
    'applyAdvFilt

    'Unload Me
    
    
    Select Case (typ)
        
        Case "Current_Active"
            Sheet2.Range("A3").Value = True
            
        Case "Current_Renewal"
            Sheet23.Range("A3").Value = True
            
        Case "Current_Amendment"
            Sheet4.Range("A3").Value = True
            
        Case "Current_Extension"
            Sheet24.Range("A3").Value = True
            
        Case "New Procurement_(No Existing)"
            Sheet3.Range("A3").Value = True
            
        Case "New Procurement_(Replace Existing)"
            Sheet5.Range("A3").Value = True
            
        Case "Closed"
            Sheet20.Range("A3").Value = True
            
        Case "Deviation"
            Sheet10.Range("A3").Value = True
            
    End Select
    
    
    
    settings.Restore
    
    Exit Sub
    
Handler:
    
    
    settings.Restore
        
    updateLog ThisWorkbook, Err.Number & ":" & Err.Description, "Save Updates Procedure: Unsuccessful"
    

End Sub


Private Sub ActBtn_addnewcontract_Click()

    Dim Security As String
    Dim cntrls As Control
    
    Security = Me.Security.Caption
    
    Unload Me
    
    If Security <> "Admin" Then
        
        For Each cntrls In EditContracts2.Controls
            
            If TypeName(cntrls) = "TextBox" Or TypeName(cntrls) = "ComboBox" Then
                
                cntrls.Enabled = False
            
            End If
        
        Next cntrls
        
        EditContracts2.Field_28.Enabled = True
        
        EditContracts2.Field_29.Enabled = True
        
        EditContracts2.Field_34.Enabled = True
    
    Else
        
        For Each cntrls In EditContracts2.Controls
            
            cntrls.Enabled = True
        
        Next cntrls
    
    End If
    
    addNewContract


End Sub


Private Sub Btn_CLMSlink_Click()
'Redirects to CLMS link
    If Field_5.Value <> "" Then
        
        ThisWorkbook.FollowHyperlink Address:="https://nedhhs.cobblestone.software/Core/ContractDetails.aspx?ID=" & Field_5.Value
    
    End If

End Sub


Private Sub Btn_Link_Click()
'Linking/Unlinking Contract To Procurement Record

    
    Dim contrNum As String
    Dim i As Long
   
   If Me.Btn_Link.Caption = "LINK CONTRACT" Then
        'Link Contract
        
            contrNum = Field_32.Value
            
            For i = 2 To Sheet8.Cells(1, 1).End(xlDown).row

                If Sheet8.Cells(i, 4).Value = contrNum Then

                     Field_36.Caption = Sheet8.Cells(i, db_DaysLeftFrRen).Value

                End If

            Next i
            
            Me.Btn_Link.Caption = "UNLINK CONTRACT"
            
    Else
        'Unlink Contract
        
        Field_32.Value = ""
        
        Me.Btn_Link.Caption = "LINK CONTRACT"

    End If
    
End Sub




Private Sub Btn_GeneralInfo_Click()

If Me.MultiPage1.Value >= 1 And Me.MultiPage1.Value < 4 Then
    
    Me.MultiPage1(0).TransitionEffect = 7
Else

    Me.MultiPage1(0).TransitionEffect = 3
End If


Me.MultiPage1.Value = 0

highlight Me.Btn_GeneralInfo
Exit_highlight Me.Btn_DeviationDetails
Exit_highlight Me.Btn_Amendments
Exit_highlight Me.Btn_RenewalInfo
Exit_highlight Me.Btn_Status
Me.ScrollHeight = 0
Me.TempmsgBox.Visible = False

End Sub



Private Sub Btn_DeviationDetails_Click()

Dim i As Long
Dim tp As Double
If Me.MultiPage1.Value >= 2 And Me.MultiPage1.Value <= 4 Then
    
    Me.MultiPage1(1).TransitionEffect = 7
Else

    Me.MultiPage1(1).TransitionEffect = 3
End If

If Me.Field_3.Value = "Deviation" Then
    
    Me.MultiPage1.Value = 1
    Exit_highlight Me.Btn_GeneralInfo
    highlight Me.Btn_DeviationDetails
    Exit_highlight Me.Btn_Amendments
    Exit_highlight Me.Btn_RenewalInfo
    Exit_highlight Me.Btn_Status
    Me.ScrollHeight = 0
    
Else
    
    Me.TempmsgBox.Visible = True

End If

 
        
End Sub



Private Sub Btn_Amendments_Click()

If Me.MultiPage1.Value >= 3 And Me.MultiPage1.Value <= 4 Then
    
    Me.MultiPage1(2).TransitionEffect = 7
Else

    Me.MultiPage1(2).TransitionEffect = 3
End If

Me.MultiPage1.Value = 2
    Exit_highlight Me.Btn_GeneralInfo
Exit_highlight Me.Btn_DeviationDetails
highlight Me.Btn_Amendments
Exit_highlight Me.Btn_RenewalInfo
Exit_highlight Me.Btn_Status
Me.ScrollHeight = 0
Me.TempmsgBox.Visible = False


End Sub




Private Sub Btn_RenewalInfo_Click()

If Me.MultiPage1.Value = 4 Then
    
    Me.MultiPage1(3).TransitionEffect = 7
Else

    Me.MultiPage1(3).TransitionEffect = 3
End If


Me.MultiPage1.Value = 3

    Exit_highlight Me.Btn_GeneralInfo
Exit_highlight Me.Btn_DeviationDetails
Exit_highlight Me.Btn_Amendments
highlight Me.Btn_RenewalInfo
Exit_highlight Me.Btn_Status
Me.TempmsgBox.Visible = False

'Me.ScrollHeight = 500

End Sub



Private Sub Btn_Status_Click()

If Me.MultiPage1.Value = 0 Then
    
    Me.MultiPage1(4).TransitionEffect = 7
Else

    Me.MultiPage1(4).TransitionEffect = 3
End If

Me.MultiPage1.Value = 4

    Exit_highlight Me.Btn_GeneralInfo
Exit_highlight Me.Btn_DeviationDetails
Exit_highlight Me.Btn_Amendments
Exit_highlight Me.Btn_RenewalInfo
highlight Me.Btn_Status
Me.ScrollHeight = 0
Me.TempmsgBox.Visible = False

'Me.ScrollHeight = 500

End Sub

Private Sub Field_3_Change()

    If Field_3.Value = "New Procurement_(Replace Existing)" Then
        
        Field_32.Visible = True
        
        Field_Header_32.Visible = True
        
        Btn_Link.Visible = True
    
    Else
        
        Field_32.Visible = False
        
        Field_Header_32.Visible = False
        
        Btn_Link.Visible = False
    
    End If
    
    If Field_3.Value = "Deviation" Then
        
        Me.Btn_DeviationDetails.Enabled = True
    Else
        
        Me.Btn_DeviationDetails.Enabled = False

    End If
    
    
    If Field_3.Value = "" Then
    
        Field_3.BorderStyle = fmBorderStyleSingle
        
        Field_3.BorderColor = ThemeUf.EmptyTB.BorderColor
    
    Else
    
        Field_3.BorderColor = ThemeUf.FilledTB.BorderColor
    
    End If

    
End Sub


Private Sub Field_11_Change()
    
    If Field_11.Value <> "" And Primary_Key.Visible = True Then
        
        Btn_AssignPco1.Visible = True
    
    Else
        
        Btn_AssignPco1.Visible = False
    
    End If

End Sub
Private Sub Field_12_Change()

    If Field_12.Value <> "" And Primary_Key.Visible = True Then
        
        Btn_AssignPco2.Visible = True
    
    Else
        
        Btn_AssignPco2.Visible = False
    
    End If

End Sub

Private Sub UserForm_Initialize()
    'Width = Application.Width * 0.95
    'Height = Application.Height * 0.95
        
    Dim cntrl As Control
    
    For Each cntrl In Me.Controls
    
        If Left(cntrl.name, 2) = "Bg" Then setProperties cntrl, ThemeUf.PageBackGround
        
        If Right(cntrl.name, 5) = "title" Then
            
            setProperties cntrl, ThemeUf.SampleInactive
                    
        End If
        
        
        If Left(cntrl.name, 3) = "Act" Then highlight cntrl
        
        If Left(cntrl.name, 3) = "Btn" Then Exit_highlight cntrl
        
        If Right(cntrl.name, 7) = "heading" Then
        
            setProperties cntrl, ThemeUf.sampleHeading
            
        End If
        
        
        If InStr(1, cntrl.name, "Field_Header", vbTextCompare) > 0 Then
        
            cntrl.ForeColor = ThemeUf.PageBackGround.ForeColor
        
        End If
        
        If TypeName(cntrl) = "ComboBox" Or TypeName(cntrl) = "TextBox" Then
        
            cntrl.FontName = Arial
            cntrl.Font.Size = 10
        
        End If
        
        If InStr(1, cntrl.name, "Field", vbTextCompare) > 0 And InStr(1, cntrl.name, "Field_Header", vbTextCompare) = 0 Then
        
            baseField cntrl
        
        End If
    Next cntrl

    
Btn_GeneralInfo_Click
    
    
    If InStr(1, Me.Field_13.Value, "Contract Executed", vbTextCompare) > 0 And Sheet12.Range("B5").Value = Me.Field_2.Caption And Left(Me.Field_3.Value, 7) <> "Current" Then
        
        Me.Field_42.Value = Now
        
    End If

    
   For Each pg In Me.MultiPage1.Pages
    
        pg.TransitionEffect = 3
        pg.TransitionPeriod = 500
   
   Next pg
    
    Me.ListBox2.ColumnCount = 2
    Me.ListBox2.ColumnWidths = "300;1"
    
    Me.ListView1.OLEDropMode = ccOLEDropManual
    
    If Me.Field_3.Value = "Deviation" Then Me.Btn_DeviationDetails.Enabled = True
    
    Me.Height = 650
    Me.Bg_Pg1.Top = 44
    Me.bg_pg2.Top = 44
    Me.bg_pg3.Top = 44
    Me.bg_pg4.Top = 44
    Me.bg_pg5.Top = 44
    
    
End Sub


Private Sub UserForm_Terminate()

    Dim cntrls As Control
    Dim Values As New Scripting.Dictionary
    Dim rowNo As Integer, colNo As Integer

    Dim settings As New ExclClsSettings
    
    Dim yesNo As String

    If Me.ActBtn_NewContract.Visible Then Exit Sub

    If Me.db_Row.Caption <> "" Then
        rowNo = CInt(Me.db_Row.Caption)
    Else
    
        settings.TurnOn
        
        Exit Sub
    
    End If
    
    For Each cntrls In Me.Controls
    
        If InStr(1, cntrls.name, "Field", vbTextCompare) > 0 And InStr(1, cntrls.name, "Field_Header", vbTextCompare) = 0 Then
            
            If TypeName(cntrls) = "Label" Then
                
                Values.Add cntrls.name, cntrls.Caption
                
            Else
                If cntrls.name = "Field_28" And Me.TempContrNum.Caption <> "" Then
                
                    Values.Add cntrls.name, Me.TempContrNum.Caption & ";" & cntrls.Value
                
                ElseIf cntrls.name = "Field_33" Then
                
                    'Values.Add cntrls.name, Me.FileLocationData.Caption
                    
                Else
                    Values.Add cntrls.name, cntrls.Value
                End If
            End If
        
        End If
    
    Next cntrls
    
    
    For colNo = 2 To Sheet8.Range("A1").End(xlToRight).Column - 3
         
    
        If colNo <> 33 And Trim(CStr(Sheet8.Cells(rowNo, colNo).Value)) <> Trim(Values("Field_" & colNo)) And (CStr(Sheet8.Cells(rowNo, colNo).Value) <> "" And Values("Field_" & colNo) <> 0) Then

            Debug.Print Me.Controls("Field_Header_" & colNo).Caption, Len(Values("Field_" & colNo)), Len(Sheet8.Cells(rowNo, colNo).Value)

            Debug.Print (CStr(Sheet8.Cells(rowNo, colNo).Value) <> "" And Values("Field_" & colNo) <> 0)
            yesNo = MsgBox("Changes were made on """ & Me.Controls("Field_Header_" & colNo).Caption & """ Field in this contract record," & vbNewLine & "Would you like to save those changes?" & vbNewLine & " Click ""Yes"" to Save, Click ""No"" to close without Saving", vbYesNo, "Save Record?")

            If yesNo = vbYes Then

                ActBtn_Save_Click
                
                'applyAdvFilt

            End If

            Exit For

        End If
        
    Next colNo
    
    settings.TurnOn
    
End Sub
