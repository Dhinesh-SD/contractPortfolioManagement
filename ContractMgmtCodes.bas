Attribute VB_Name = "ContractMgmtCodes"
Option Explicit

Public Enum PageContractsLocation

 ''''''''''''''''''''''pg_TableFirstCol = 5
    pg_TableFirstrow = 18

End Enum

Sub turnonevents()
    Dim settings As New ExclClsSettings
    
    settings.TurnOn

End Sub
Sub openContractEdit()


    Dim settings As New ExclClsSettings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Turn off excel Functionality to speedup the procedure
    
    settings.TurnOn
    
    settings.TurnOff
    
 ''''''''''''''''''''''Procedure to read the corresponding Contract Data and Display it in Userform
    
 ''''''''''''''''''''''This macro is assigned to Edit Contracts button found on all pages
    
    Dim i               As Long
    Dim j               As Long
    Dim selectedRow     As Long
    Dim pk              As String
    Dim ws              As Worksheet
    Dim pkCol           As Long
    Dim cntrls          As Control
    Dim columnNum       As Long
    Dim ted             As Date '''''''''''''''''''''''term End Date
    Dim Med             As Date ''''''''''''''''''''''' Max end date
    Dim Nor             As Integer ''''''''''''''''''''''' No. Of Renewals
    Dim Nrd             As Date '''''''''''''''''''''''Next renewal dates
    Dim Erd             As Double ''''''''''''''''''''''' each renewal duration
    Dim Tsd             As Date '''''''''''''''''''''''Term start date
    Dim Tdy             As Date
    Dim Ext             As Double
    Dim years           As Double
    Dim sh              As Worksheet
    Dim tempLabel()     As String
    Dim pg_TableFirstCol As Integer
    Set ws = ActiveSheet
                    
 ''''''''''''''''''''''Sheet8 is the database sheet
    
    Set sh = Sheet8
                    
 ''''''''''''''''''''''Find and Set Primary key Column as its column location may change from page to page
    
    For i = 4 To ws.Cells(17, 4).End(xlToRight).Column
                    
        If ws.Cells(17, i).Value = "Primary_Key" Then
            
            pkCol = i
        
        ElseIf ws.Cells(17, i).Value = "Priority" Then
        
            pg_TableFirstCol = i
            
        End If
        
    
    Next i
            
 ''''''''''''''''''''''Set the list of inputs for Linked Contracts field
    With EditContracts2
        
        .Field_32.RowSource = "Sheet1!" & Range("ContractNums#").Address
                
 ''''''''''''''''''''''Set style of Field_32 ComboBox
        
        .Field_32.Style = 2
                
 ''''''''''''''''''''''If the page contains no contracts list, i.e if the first entry of the table is empty( check if primary key is blank) then call addNew record  instead of
        
        If Sheet16.Cells(pg_TableFirstrow, pg_TableFirstCol).Value = "" Or ws.Cells(pg_TableFirstrow, pkCol).Value = "" Then
                
 ''''''''''''''''''''''Call Add New Contract
            
            addNewContract
            
            settings.Restore
            
            Exit Sub
        
        End If
                
 ''''''''''''''''''''''Conditions if the user clicks the edit contracts button while selecting cells not in the contracts list range Exit sub
        
        If Sheet16.Cells(pg_TableFirstrow, pg_TableFirstCol).Value = "" Or ActiveCell.row > ws.Cells(pg_TableFirstrow - 1, pg_TableFirstCol).End(xlDown).row Or ActiveCell.row < pg_TableFirstrow Then
            
            settings.Restore
            
            Exit Sub
        
        End If
                
 '''''''''''''''''''''' Set primary key value to pk
        
        pk = ws.Cells(ActiveCell.row, pkCol).Value
        
 ''''''''''''''''''''''On Error Resume Next
                
 ''''''''''''''''''''''Loop through all contracts in database to find matching record with primary key
        
        For i = 2 To sh.Cells(1, 1).End(xlDown).row
                
 ''''''''''''''''''''''If matching Key is found Continue reading Data
            
            If sh.Cells(i, 1).Value = pk Then
                
                .db_Row.Caption = i
                
 ''''''''''''''''''''''Fill Label Primary_key in  Userform
                
                .Primary_Key.Caption = pk
                
 ''''''''''''''''''''''Loop through all the controls in editcontrols userform
                
                For Each cntrls In .Controls
                
 ''''''''''''''''''''''If the control Name contains "Field" then it needs to be populated with the corresponding value from the database
                    
                    If (TypeName(cntrls) = "TextBox" Or TypeName(cntrls) = "ComboBox") And Left(cntrls.name, 5) = "Field" Then
                
 ''''''''''''''''''''''If the type of control is "TextBox" or "ComboBox" we need to invoke the "Value" Property
                        
                        columnNum = Val(Mid(cntrls.name, InStrRev(cntrls.name, "_") + 1))
                        
                        
 ''''''''''''''''''''''Debug.Print cntrls.Value
                        
                        
                        On Error Resume Next
                        
                        cntrls.Value = sh.Cells(i, columnNum).Value
                        
                        On Error GoTo 0
                        
                        If columnNum = 33 Then
                        
                            cntrls.Value = vbNullString
                            
                        End If
                    
                    ElseIf TypeName(cntrls) = "Label" And Left(cntrls.name, 5) = "Field" And InStr(1, cntrls.name, "Field_Header", vbTextCompare) = 0 Then
                
 ''''''''''''''''''''''If the control is a label then we need to invoke the ".Caption" property to populate this field
                        
                        columnNum = Val(Mid(cntrls.name, InStrRev(cntrls.name, "_") + 1))
                        
                        cntrls.Caption = sh.Cells(i, columnNum).Value
                        
                        If columnNum = 26 And sh.Cells(i, columnNum).Value < 1 And cntrls.Caption <> "" Then
                
 ''''''''''''''''''''''Conditions for Renewal Duration(Field_26), if renewal duration is less than 1 year Mark the years comboBox as zero and Monts combobox with the repective number of months
                            
                            If cntrls.Caption = "" Then cntrls.Caption = 0
                
 ''''''''''''''''''''''Set the years value from the renewal duraion
                            
                            years = CDbl(cntrls.Caption)
                            
                            .ComboBox1.Value = 0
                
 ''''''''''''''''''''''Calculations to calculate the durations in months for future calculations( to calculate the term dates)
                            
                            .ComboBox2.Value = sh.Cells(i, columnNum).Value * 12
                        
                        ElseIf columnNum = 26 And sh.Cells(i, columnNum).Value >= 1 Then
                
 ''''''''''''''''''''''If renewal Duration is greater than 1 year then  follow the below procedure
                            
                            If cntrls.Caption = "" Then cntrls.Caption = 0
                            
                            years = CDbl(cntrls.Caption)
                            
                            .ComboBox1.Value = Application.WorksheetFunction.RoundDown(sh.Cells(i, columnNum).Value, 0)
                            
                            .ComboBox2.Value = (sh.Cells(i, columnNum).Value - Application.WorksheetFunction.RoundDown(sh.Cells(i, columnNum).Value, 0)) * 12
                
 ''''''''''''''''''''''Similar conditions for Extension Duration(Field_30)
                        
                        ElseIf columnNum = 30 And sh.Cells(i, columnNum).Value < 1 Then
                
 ''''''''''''''''''''''Extension Duration less than 1 year
                            
                            If cntrls.Caption = "" Then cntrls.Caption = 0
                            
                            years = CDbl(cntrls.Caption)
                            
                            .ComboBox3.Value = 0
                            
                            .ComboBox4.Value = sh.Cells(i, columnNum).Value * 12
                        
                        ElseIf columnNum = 30 And sh.Cells(i, columnNum).Value >= 1 Then
                 
 ''''''''''''''''''''''Extension Duration Greater than or equal to 1 year
                            
                            If cntrls.Caption = "" Then cntrls.Caption = 0
                            
                            years = CDbl(cntrls.Caption)
                            
                            .ComboBox3.Value = Application.WorksheetFunction.RoundDown(sh.Cells(i, columnNum).Value, 0)
                            
                            .ComboBox4.Value = (sh.Cells(i, columnNum).Value - Application.WorksheetFunction.RoundDown(sh.Cells(i, columnNum).Value, 0)) * 12
                        
                        End If
                    
                    End If
                
                Next cntrls
                
            Dim strfile As String, subStrFile As String, row As Integer
            
            Dim folderName As String
            Dim FSOLibrary As Object
            Dim FSOFolder As Object
            Dim FSOFile As Object
            Dim folderEmpty As Boolean
            
            folderEmpty = True
            
            row = 0
            
            Dim Obj As Object, rootFolder As String, contractFolder

            rootFolder = Replace(ThisWorkbook.FullName, ThisWorkbook.name, "PCO Contract Files")
            
            Set Obj = CreateObject("Scripting.fileSystemObject")
            
            If Not Obj.FolderExists(rootFolder) Then
            
                Obj.CreateFolder (rootFolder)
            
            End If
                    
            contractFolder = rootFolder & "\" & .Primary_Key.Caption & " " & .Field_8.Value
            
            strfile = Dir(rootFolder & "\")
            
                       
            Set FSOLibrary = CreateObject("Scripting.FileSystemObject")
            Set FSOFolder = FSOLibrary.GetFolder(rootFolder & "\")
            
             
                 For Each FSOFile In FSOFolder.subFolders
                
                    If .Field_4.Value = Left(FSOFile.name, Len(.Field_4.Value)) And .Field_4.Value <> "" Then
                    
                        If Obj.FolderExists(contractFolder) Then
                        
                            subStrFile = Dir(rootFolder & "\" & FSOFile.name & "\")
                            
                            Do While Len(subStrFile) > 0
                            
 ''''''''''''''''''''''Debug.Print FSOFile.name
                                
                                If Not Obj.FileExists(contractFolder & "\" & subStrFile) Then
                                
                                    Name rootFolder & "\" & FSOFile.name & "\" & subStrFile As contractFolder & "\" & subStrFile
                                Else
                                
                                folderEmpty = False
                                
                                End If
                                
                                subStrFile = Dir
                                
                            Loop
                            
 ''''''''''''''''''''''Debug.Print FSOFile.name
                        
                        If folderEmpty Then RmDir rootFolder & "\" & FSOFile.name & "\"
                        
                        Else
                        
                            Name rootFolder & "\" & FSOFile.name & "\" As contractFolder & "\"
                        
                        End If
                    
                    End If
                    
                Next
 ''''''''''''''''''''''Release the memory
            
            .Field_33.Value = contractFolder
            
            Set FSOLibrary = Nothing
            Set FSOFolder = Nothing
            
            
                
            If Not Obj.FolderExists(Replace(contractFolder, "/", "")) Then
                
                Obj.CreateFolder (Replace(contractFolder, "/", ""))
            
            End If
                        
            contractFolder = Replace(contractFolder, "/", "")
                     
            If contractFolder <> "" And Obj.FolderExists(contractFolder & "\") Then
            
 ''''''''''''''''''''''.Btn_ChangeDirectory.Visible = True
                
                strfile = Dir(contractFolder & "\")
                
                Do While Len(strfile) > 0
                    .ListBox2.AddItem
                    .ListBox2.List(row, 0) = strfile
                    .ListBox2.List(row, 1) = contractFolder & "\"
                    row = row + 1
                    strfile = Dir
                Loop
                                
            End If
                        
            Set Obj = Nothing
                        
            Tsd = sh.Cells(i, db_startDate).Value
            
            ted = sh.Cells(i, db_EndDate).Value
            
            Med = sh.Cells(i, db_MaxendDate).Value
            
            Nor = sh.Cells(i, db_NoOfRens).Value
            
            Erd = sh.Cells(i, db_EachRenDur).Value
            
            Ext = sh.Cells(i, db_ExtensionDur).Value
            
            Tdy = Now
            
 ''''''''''''''''''''''User Defined Function to calculate the term Start Dates and end dates to display in the Edit contracts form
            
            CalcTerm Tsd, ted, Nor, Erd, Ext
            
            Exit For
            
            End If
        
        Next i
                
 ''''''''''''''''''''''Condition to display Assign Temp Pco, If Temporary Pco-1 is already assigned to a contract The assign temp Pco-1 button will not be visible
        
        If .Field_11.Value <> "" Then .Btn_AssignPco1.Visible = False
                
 ''''''''''''''''''''''Condition to display Assign Temp Pco, If Temporary Pco-2 is already assigned to a contract The assign temp Pco-2 button will not be visible
        
        If .Field_12.Value <> "" Then .Btn_AssignPco2.Visible = False
                
 ''''''''''''''''''''''Fill the Security Label with Assigned security for the user from profile Information sheet ("Admin/User")
        
        .Security.Caption = Sheet12.Range("Security").Value
                
 ''''''''''''''''''''''Set Combobox style to 2 :- this will disable users from entering arbitrary values and can only choose values available in the combobox list
        
        For Each cntrls In .Controls
            
            If TypeName(cntrls) = "ComboBox" Then
                
                cntrls.Style = 2
            
            End If
        
        Next cntrls
        
 ''''''''''''''''''''''Conditions to display Link Contract button or unlink contract button
        
        If .Field_32.Value <> "" Then
                
 ''''''''''''''''''''''If this record has a contract already linked to it hen hide link contract and display unlink contract button
            
            .Btn_Link.Caption = "UNLINK CONTRACT"
            
        
        End If
        
 ''''''''''''''''''''''If account is not with Admin Security then
        
        If .Security.Caption <> "Admin" Then
                
 ''''''''''''''''''''''Disable controls to lock users from making changes
            
            For Each cntrls In .Controls
                
                If TypeName(cntrls) = "TextBox" Or TypeName(cntrls) = "ComboBox" Then
                    
                    cntrls.Enabled = False
                
                End If
            
            Next cntrls
                    
            .ActBtn_AddNewContract.Visible = False
                
 ''''''''''''''''''''''Enable users to just edit few fields if .enabled = true then users can edit those fields.
            
             
            .Field_5.Enabled = True
            
            .Field_19.Enabled = True
            
            .Field_20.Enabled = True
            
            .Field_21.Enabled = True
            
            .Field_28.Enabled = True
            
            .Field_29.Enabled = True
            
            .Field_31.Enabled = True
            
            .Field_34.Enabled = True
            
            .Field_39.Enabled = True
            
            .Field_40.Enabled = True
            
 ''''''''''''''''''''''If the selected contract is a temporary assignment to another PCO then the remarks field will be disabled. as it contains
            
            If Left(.Primary_Key.Caption, 4) = "Temp" Then
                
                .TempContrNum.Visible = True
                
                tempLabel = Split(.Field_28.Value, ";")
                .Field_28.Value = tempLabel(1)
                .TempContrNum.Caption = tempLabel(0)
                
            Else
            .TempContrNum.Visible = False
            
            
            End If
        Else
        
 ''''''''''''''''''''''conditions for accounts with admin access
            
            For Each cntrls In .Controls
                
                cntrls.Enabled = True
            
            Next cntrls
                
 ''''''''''''''''''''''If the selected record is a temporary assignment the primary key will have a Temp keyword assigned to it
            
            If Left(pk, 4) = "Temp" Then
                
 ''''''''''''''''''''''The following fields and buttons will be disabled for temporarily assigned contracts
                
                .Field_11.Enabled = False
                
 ''''''''''''''''''''''Temporary assignment of PCO cannot be made for a record marked as a temporary
                
                .Field_12.Enabled = False
                
                .Btn_UnassignPco1.Visible = False
                
                .Btn_UnassignPco2.Visible = False
            
            Else
                
 ''''''''''''''''''''''If the contract has a temporary assignment Display unassign Temporary Pco Button and disable the temporary Assignment field
                
                If .Field_11.Value <> "" Then
                
                    .Field_11.Enabled = False
                    
                    .Btn_UnassignPco1.Visible = True
                                
                End If
                
                If .Field_12.Value <> "" Then
                    
                    .Field_12.Enabled = False
                    
                    .Btn_UnassignPco2.Visible = True
                
                End If
            
            End If
                
 ''''''''''''''''''''''Disable Remarks field for temporary assgnment as it has the contract for which this record was a temporary assignment and we dont want any changes to be made in that field!
            
              If Left(.Primary_Key.Caption, 4) = "Temp" Then
                
                
                tempLabel = Split(.Field_28.Value, ";")
                .Field_28.Value = tempLabel(1)
                .TempContrNum.Caption = tempLabel(0)
            End If
        
        End If
        
 ''''''''''''''''''''''Condition if status is "Contract Executed"
        
        If .Field_13.Value = "Contract Executed (Final Status)" And Left(Sheet12.Range("position"), 3) = "PCO" Then
 ''''''''''''''''''''''Enable all fields for PCO'''''''''''''''''''''''s to make changes if the status is marked as contract executed!
            For Each cntrls In .Controls
                
                cntrls.Enabled = True
            
            Next cntrls
        
        .Field_13.Enabled = False
        
        .Field_34.Value = "High"
        
        .Field_34.Enabled = False
        
        .MSG.Caption = "Contract Has been Executed! Please fill All the Required Details and Save It as a Current_Active Contract"
        
        End If
 ''''''''''''''''''''''Show the userform after all the data is read and populated in their respective fields
        
            
            .Top = Application.Top + Application.Height / 2 - .Height / 2
            
            .Left = Application.Left + Application.Width / 2 - .Width / 2
        
        .MultiPage1.Value = 0
 ''''''''''''''''''''''If CLMS number is populated Display view contract Button(Hyperlink to CLMS website)
        If .Field_5.Value <> "" Then .Btn_CLMSlink.Visible = True
        
        .ActBtn_NewContract.Visible = False
        
        Dim yesNo As String
        
        .Field_36.Caption = IIf(.Field_36.Caption = "", 0, .Field_36.Caption)
        
        If .Field_43.Value = "" And CInt(.Field_36.Caption) < 90 And .Field_36.Caption <> 0 Then
        
            yesNo = MsgBox("This contract is nearing renewal date, Would you like to mark this for Renewal?", vbYesNo, "Mark FOr Renewal?")
                
            If yesNo = vbYes Then
                .Field_43.Value = "Yes"
                                
                .Field_3.Value = "Current_Renewal"
                
                .changesMade.Caption = "True"
                
            Else
                .Field_43.Value = "No"
                                                
                .changesMade.Caption = "True"
                
            End If
        
        End If
         .Field_33.Enabled = False
         
         If .Field_34.Value <> "" Then .Priority.Caption = "  " & .Field_34.Value & " Priority Contract"
        
        .Show vbModeless
        
    End With
    
    settings.Restore


End Sub


Sub addNewContract()

    Dim settings As New ExclClsSettings
    Dim cntrl As Control
    
 '''''''''''''''''''''''Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    settings.TurnOff
 ''''''''''''''''''''''Procedure to Add a new contract to the database
 ''''''''''''''''''''''This is an Admin Level procedure
    If Sheet12.Range("Security") <> "Admin" Then
        
        settings.Restore
        
        Exit Sub
        
    End If
 ''''''''''''''''''''''Hide all the buttons which are visible in edit contracts userform and display buttons to add new contract
    
    With EditContracts2
            
    
    For Each cntrl In .Controls
    
        If Right(cntrl.name, 5) = "title" Then
            
            cntrl.Caption = "ADD NEW CONTRACT"
            
        End If
        
    Next cntrl

        .ActBtn_Save.Visible = False
        
        .ActBtn_NewContract.Visible = True
        
        .ActBtn_AddNewContract.Visible = False
            
        .Primary_Key.Visible = False
        
        .PCO.Visible = True
        
        .Btn_Next.Visible = False
        
        .Btn_Previous.Visible = False
        
        .Field_32.RowSource = "Sheet1!" & Range("ContractNums#").Address
        
        .Field_32.Style = 2
    
        .Field_11.Enabled = False
        
        .Field_12.Enabled = False
        
        .ActBtn_DeleteContract.Visible = False
        
        .Top = Application.Top + Application.Height / 2 - .Height / 2
        
        .Left = Application.Left + Application.Width / 2 - .Width / 2
    
        .MultiPage1.Value = 0
                
        .Show vbModeless
    
    End With

    settings.Restore
End Sub





Sub CalcTerm(Tsd As Date, ted As Date, Nor As Integer, Erd As Double, Ext As Double)
    
    Dim settings As New ExclClsSettings
    '''''''''''''''''''''''Turn off excel Functionality to speedup the procedure
    settings.TurnOn
    settings.TurnOff
    '''''''''''''''''''''''Procedure to Calculate the Terms start Date and end date for the contract Based on the Date Values provided
    EditContracts2.ListBox1.RowSource = ""
    
    Dim j   As Long
    '''''''''''''''''''''''This procedure will paste its resut in Sheet17 and the Listbox in editcontracts userform under renewal information will have this result range as its Source of data
    Sheet17.Range("AB3").CurrentRegion.Offset(2).ClearContents
    '''''''''''''''''''''''Resize result table based on the size of the result
    Sheet17.ListObjects("Table14").Resize Range("AB3:AE4")
    '''''''''''''''''''''''Populate basic data
    Sheet17.Cells(4, 29).Value = Tsd
    
    Sheet17.Cells(4, 30).Value = ted
    
    '''''''''''''''''''''''If Nor > 0 Then
        
        For j = 4 To Nor + 4
            
            If j > 4 Then
    '''''''''''''''''''''''Start populating from second row of the table as first row contains start and end dates for calculation
                Sheet17.Cells(j, 28).Value = "RENEWAL-" & j - 4
                
                Sheet17.Cells(j, 29).Value = Sheet17.Cells(j - 1, 30).Value + 1
                
                Sheet17.Cells(j, 30).Value = DateSerial(Year(Sheet17.Cells(j - 1, 30).Value) + Erd, Month(Sheet17.Cells(j - 1, 30).Value), Day(Sheet17.Cells(j - 1, 30).Value))
            
            End If
    '''''''''''''''''''''''Conditions to choose the status of the term based on dates
            If Now > Sheet17.Cells(j, 30).Value Then
                
                Sheet17.Cells(j, 31).Value = "Term Completed"
            
            ElseIf Now < Sheet17.Cells(j, 30).Value And Now > Sheet17.Cells(j, 29).Value Then
                
                Sheet17.Cells(j, 31).Value = "Current Term<<"
            
            ElseIf Now < Sheet17.Cells(j, 29).Value Then
                
                Sheet17.Cells(j, 31).Value = "Yet to Start"
            
            End If
        
        Next j
   
   ''''''''''''''''''''''' End If
    '''''''''''''''''''''''If the contract has extension duration >0 then this procedure will take place similar to the renewals
    If Ext <> 0 Then
        
        j = Nor + 4
        
        Sheet17.Cells(j + 1, 28).Value = "Extension"
        
        Sheet17.Cells(j + 1, 29).Value = Sheet17.Cells(j, 30).Value + 1
        
        Sheet17.Cells(j + 1, 30).Value = Sheet17.Cells(j, 30).Value + Ext * 365
        
        If Now > Sheet17.Cells(j + 1, 30).Value Then
            
            Sheet17.Cells(j, 31).Value = "Term Completed"
        
        ElseIf Now < Sheet17.Cells(j + 1, 30).Value And Now > Sheet17.Cells(j + 1, 29).Value Then
            
            Sheet17.Cells(j + 1, 31).Value = "Current Term<<"
        
        ElseIf Now < Sheet17.Cells(j, 29).Value Then
            
            Sheet17.Cells(j + 1, 31).Value = "Yet to Start"
        
        End If

    End If

    settings.Restore

End Sub


