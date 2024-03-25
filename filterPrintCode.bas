Attribute VB_Name = "filterPrintCode"
Sub filtPrint(ws As Worksheet, target As Range)

Dim rng As Range

Set rng = ws.Cells(16, target.Column)

        If target.Value = "" Then
             
             rng.Value = ""
        
        ElseIf target.Value = "Equals" Or target.Value = "" Then
            
            rng.Value = "=" & Chr(34) & "=Txt" & Chr(34)
        
        ElseIf target.Value = "Does Not Equals" Then
            
            rng.Value = "=" & Chr(34) & "<>" & Chr(34)
        
        ElseIf target.Value = "Contains" Then
             
             rng.Value = "=" & Chr(34) & "=*" & "Txt" & "*" & Chr(34)
        
        ElseIf target.Value = "Does Not Contains" Then
            
            rng.Value = "<>*" & "Txt" & "*"
        
        ElseIf target.Value = "Begins With" Then
            
            rng.Value = "=" & Chr(34) & "=" & "Txt" & "*" & Chr(34)
        
        ElseIf target.Value = "Ends With" Then
             
             rng.Value = "=" & Chr(34) & "=*" & "Txt" & Chr(34)
        
        ElseIf target.Value = "Greater Than" Or target.Value = "After" Then
            
            rng.Value = ">" & "Txt"
        
        ElseIf target.Value = "Greater Than or equal to" Then
             
             rng.Value = ">=" & "Txt"
        
        ElseIf target.Value = "Less Than or equal to" Then
            
            rng.Value = "<=" & "Txt"
        
        ElseIf target.Value = "Less Than" Or target.Value = "Before" Then
             
             rng.Value = "<" & "Txt"
        
        ElseIf target.Value = "Between" Then
             
             rng.Value = ">=" & "Txt" & ",<=" & "Txt"
        
        End If

End Sub
