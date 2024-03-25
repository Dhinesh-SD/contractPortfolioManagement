VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExclClsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private calculation As XlCalculation
Private displayStatus As Boolean
Private enableEvents As Boolean
Private screenUpdating As Boolean

'Procedure: Backup the initial Functionality
'Purpose: Backup Current Excel Settings

Public Sub Backup()

    calculation = Application.calculation
    
    displayStatus = Application.DisplayStatusBar
    
    enableEvents = Application.enableEvents
    
    screenUpdating = Application.screenUpdating

End Sub


'Procedure: Restore Functionality
'Purpose: Restore Excel Settings

Public Sub Restore()

    Application.calculation = calculation
    
    Application.DisplayStatusBar = displayStatus
    
    Application.enableEvents = enableEvents
    
    Application.screenUpdating = screenUpdating

End Sub

'Procedure: turn-Off any active Functionality
'Purpose: To turn-off Excel Settings

Public Sub TurnOff()
    Call Backup
    
    Application.calculation = xlCalculationManual
    
    Application.DisplayStatusBar = False
    
    Application.enableEvents = False
    
    Application.screenUpdating = False

    
End Sub


Public Sub TurnOn()
    
    Application.calculation = xlCalculationAutomatic
    
    Application.DisplayStatusBar = True
    
    Application.enableEvents = True
    
    Application.screenUpdating = True

    
End Sub

