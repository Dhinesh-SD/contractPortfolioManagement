VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "statClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StatusField As String
Dim statusbox As New Status

Public Sub showStatus(stat As String)
    
    statusbox.Field_Header_1.Caption = stat
    DoEvents
    statusbox.Show vbModeless
    
End Sub


Public Sub closeStats()
    
    Unload statusbox

End Sub


