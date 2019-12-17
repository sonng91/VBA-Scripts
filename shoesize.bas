Attribute VB_Name = "shoesize"
Sub shoes()
'Find the last used column in a Row: row 1 in this example
    Dim lastCol As Integer
    With Sheets(1)
        lastCol = .Cells(1, .columns.Count).End(xlToLeft).Column
    End With
    
    Dim lastRow As Long
        lastRow = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        
      
    Dim add1 As Integer
    Dim add2 As Integer
    Dim duh As String
    
    duh = " - "
    
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = "Results"
    
    x1 = 0
    x2 = 0
    y1 = 0
    y2 = 0
    
    i = 5
    x = 2
    row1 = 0
    Header = 0
    
'Sheets(1).columns("A:D").Copy
'Sheets(2).columns("A:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
Sheets("Results").Range("A1").Offset(x1, y1).Value = Sheets(1).Range("A1").Offset(x2, y2).Value
Sheets("Results").Range("B1").Offset(x1, y1).Value = Sheets(1).Range("B1").Offset(x2, y2).Value
Sheets("Results").Range("C1").Offset(x1, y1).Value = Sheets(1).Range("C1").Offset(x2, y2).Value
Sheets("Results").Range("D1").Offset(x1, y1).Value = Sheets(1).Range("D1").Offset(x2, y2).Value



While x < lastRow
    
    i = 5
    add2 = 0
    add1 = 0
    x1 = row1 + 1
    x2 = x2 + 1
        Sheets(2).Range("A1").Offset(x1, y1).Value = Sheets(1).Range("A1").Offset(x2, 0).Value
        Sheets(2).Range("B1").Offset(x1, y1).Value = Sheets(1).Range("B1").Offset(x2, 0).Value
        Sheets(2).Range("C1").Offset(x1, y1).Value = Sheets(1).Range("C1").Offset(x2, 0).Value
        Sheets(2).Range("D1").Offset(x1, y1).Value = Sheets(1).Range("D1").Offset(x2, 0).Value
    
    While i < lastCol + 1

        
        result = Sheets(1).Range("E1").Offset(, add1).Value & duh & Sheets(1).Range("E2").Offset(r1, add1).Value
        
        If Sheets(1).Range("E2").Offset(r1, add1) > 0 Then
        
            'Range("K2").Offset(row1, add2).NumberFormat = "@"
            'Range("K2").Offset(row1, add2).Value = result
            'Sheets(2).Range("AM2").Offset(row1, add2).NumberFormat = "@"
            'Sheets(2).Range("AM2").Offset(row1, add2).Value = result
            
            Sheets(2).Range("E2").Offset(row1, add2).NumberFormat = "@"
            Sheets(2).Range("E2").Offset(row1, add2).Value = result
            
            
            add1 = add1 + 1
            'add2 = add2 + 1
            row1 = row1 + 1
            
            GoTo step2
            
        Else
            add1 = add1 + 1
            GoTo step2
            
        End If
        
step2:
    
        i = i + 1
        
        
        
        Wend
    r1 = r1 + 1
    x = x + 1

Wend
    
End Sub
