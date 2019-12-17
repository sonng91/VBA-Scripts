Attribute VB_Name = "delim"
Sub delimiter()

Dim Timer As Date
Dim x As Range
Dim toRemove()
Dim Memory As String
Dim lastRow As Long

Timer = Now()
previousVar = 0
lastRow = Range("B" & Rows.Count).End(xlUp).Row

t = 0
focusedCell = 2
toRemove() = Array("_", "/", "(", "-", " ", ",", ".", ":", "|||", "||")


While focusedCell < Int(lastRow)

    For focusedCell = focusedCell To Int(lastRow)
    
        For Each itm In toRemove()
                
                Range("A" & focusedCell).Value = Replace(Range("A" & focusedCell), itm, "|")
                Memory = Range("A" & focusedCell).Value
        
        Next itm
        
        y = Split(Memory, "|")
        
        For i = 0 To UBound(y)
        
            If IsNumeric(y(i)) = True Then
                
                If Len(y(i)) = 9 And y(i) > 100000000 And y(i) <> previousVar Then
                    
                    Tracking = Range("B" & focusedCell).Value 'Copy tracking number cell
                    Shipper = Range("C" & focusedCell).Value
                    Range("D" & focusedCell).Offset(t, 0).Value = y(i)
                    Range("E" & focusedCell).Offset(t, 0).NumberFormat = "@"
                    Range("E" & focusedCell).Offset(t, 0) = Tracking
                    
                    If Range("C1").Value <> "" Then
                        Range("D" & focusedCell).Offset(t, 2) = Shipper
                    End If
                    
                    t = t + 1
                    previousVar = y(i)
                End If
        
            End If
        Next i
        t = t - 1
    Next focusedCell

Wend

    ActiveSheet.Range("D:F").RemoveDuplicates columns:=Array(1, 2), Header _
        :=xlNo
        
    columns("A:C").Delete
    
Range("A1").Value = "External Order ID"
Range("B1").Value = "Tracking Number"

If Range("C2").Value <> "" Then
    Range("C1").Value = "Shipper Name"
End If

MsgBox "Ellapsed Time in Hrs:Min:Sec :" & Format(Now() - Timer, "hh:mm:ss")

End Sub
