Attribute VB_Name = "QVC"
Sub QVCmacro1()

Dim i As Integer
Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
Dim str2 As String
Dim name2 As String

Dim c As Object
Dim tmp As String


name2 = Range("B2").Value

Worksheets.Add(After:=Worksheets(Worksheets.Count)).name = name2
 
Sheets(name2).Range("A1").Value = "Control Nbr"
Sheets(name2).Range("B1").Value = "PO Number"
Sheets(name2).Range("C1").Value = "SKU"
Sheets(name2).Range("D1").Value = "zulily Qty sold"
Sheets(name2).Range("E1").Value = "Contracted Cost 5 + 2 decimals"
Sheets(name2).Range("F1").Value = "Ship To Identifier"

Sheets(name2).Range("A1").Font.Bold = True
Sheets(name2).Range("B1").Font.Bold = True
Sheets(name2).Range("C1").Font.Bold = True
Sheets(name2).Range("D1").Font.Bold = True
Sheets(name2).Range("E1").Font.Bold = True
Sheets(name2).Range("F1").Font.Bold = True

i = 2
j = 0


While i < lastRow + 1
    
    'Control Nbr
    'Sheets(name2).Range("A" & i).NumberFormat = "0000000000"
    Sheets(name2).Range("A" & i).NumberFormat = "@"
    If Len(Sheets(1).Range("A" & i)) < 11 Then
        tmp = Format(Sheets(1).Range("A" & i).Value, "0000000000")
        Sheets(name2).Range("A" & i).Value = tmp
    End If
    
    
    
    str2 = Sheets(1).Range("B" & i).Value
    Dim POname As String
    Dim POname1 As String
    
    'PO Number column
    'Sheets(name2).Range("B" & i).Value = Mid(str2, 1, 10)
    POname = Sheets(1).Range("B" & i).Value
    POname1 = Replace(POname, "-", "")
    Sheets(name2).Range("B" & i).Value = POname1
    
    
    'SKU column
    Sheets(name2).Range("C" & i).Value = Sheets(1).Range("E" & i)
    
    'zulily Qty sold column
    Sheets(name2).Range("D" & i).NumberFormat = "@"
    If Len(Sheets(1).Range("H" & i)) < 11 Then
        tmp = Format(Sheets(1).Range("H" & i).Value, "0000000000")
        Sheets(name2).Range("D" & i).Value = tmp
    End If
    'Range(Cells(i, 4), Cells(i, 5)).Merge
    
    'Contracted Cost 5 + 2 decimals
    Sheets(name2).Range("E" & i).NumberFormat = "@"
        cost = Sheets(1).Range("M" & i).Value
        cost2 = cost * 100
        If Len(cost2) < 8 Then
            tmp = Format(cost2, "0000000")
            Sheets(name2).Range("E" & i).Value = tmp
        End If
    
    
    
    'Ship To Identifier
    If Mid(str2, 12, 2) = 12 Then
        Sheets(name2).Range("F" & i).NumberFormat = "@"
        Sheets(name2).Range("F" & i).Value = "00000000000000000012"
        
    ElseIf Mid(str2, 12, 1) = 8 Then
        Sheets(name2).Range("F" & i).NumberFormat = "@"
        Sheets(name2).Range("F" & i).Value = "00000000000000000008"
        
    Else
        A = InputBox("What Dist Center is this? " + str2)
        Sheets(name2).Range("F" & i).Value = A
        
    End If
    
    i = i + 1

Wend

'Range(Cells(1, 4), Cells(1, 5)).Merge
columns("A:A").ColumnWidth = 20
columns("B:B").ColumnWidth = 24
columns("C:C").ColumnWidth = 20
columns("D:D").ColumnWidth = 15
columns("E:E").ColumnWidth = 30
columns("F:F").ColumnWidth = 24



End Sub


