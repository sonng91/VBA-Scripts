Attribute VB_Name = "AddingWS"
Sub AddWorksheets()

Dim AddWS() As Variant
Dim lastRow As Long
    lastRow = Sheets(1).Range("E" & Rows.Count).End(xlUp).Row

AddWS() = Array()

'Go through the shipper name and add it to a array if it's unique
For i = 2 To lastRow


        If Range("I" & i).Value <> AddWS() Then
            
        Else
            AddWS(itm + 1) = Range("I" & i).Value
            msg = msg & AddWS(itm) & vbNewLine
        End If
    
Next i

MsgBox "the values of my dynamic array are: " & vbNewLine & msg

End Sub
