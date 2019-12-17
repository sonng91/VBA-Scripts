Attribute VB_Name = "Zipcoding"
Sub zipcode()

Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
Dim i As Integer
Dim x As String



i = 1

While i < lastRow + 1
x = Range("P" & i).Value
        If Len(x) < 5 Then
            Range("P" & i).NumberFormat = "@"
            tmp = Format(Range("P" & i).Value, "00000")
            Range("P" & i).Value = tmp
        End If
    
    i = i + 1
    
Wend



End Sub
