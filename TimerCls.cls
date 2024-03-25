VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tim As Double

Public Sub start()
    
    tim = MicroTimer

End Sub

Public Sub PrintTime(Procedure As String)
    
    Debug.Print Round(MicroTimer - tim, 5), Procedure
    
    tim = MicroTimer

End Sub
