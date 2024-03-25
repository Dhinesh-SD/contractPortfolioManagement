VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dateConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ConvertDates(ByVal dates As String) As Date
    Dim elements() As String
    Dim yr As Integer, mon As Integer, dy As Integer
    
    elements = Split(dates, " ")
    
    yr = CInt(elements(3))
    
    dy = CInt(elements(2))
    
    Select Case elements(1)
        Case "Jan"
            mon = 1
        Case "Feb"
            mon = 2
        Case "Mar"
            mon = 3
        Case "Apr"
            mon = 4
        Case "May"
            mon = 5
        Case "Jun"
            mon = 6
        Case "Jul"
            mon = 7
        Case "Aug"
            mon = 8
        Case "Sep"
            mon = 9
        Case "Oct"
            mon = 10
        Case "Nov"
            mon = 11
        Case "Dec"
            mon = 12
    End Select
        
        
    ConvertDates = Evaluate("Date(" & yr & "," & mon & "," & dy & ")")

End Function


