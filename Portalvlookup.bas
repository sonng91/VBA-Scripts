Attribute VB_Name = "Portalvlookup"
Sub PortalvlookupStep()

Dim i As Integer

Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

    For i = 2 To lastRow
        Range("O" & i).Formula = "=IF(ISNA(VLOOKUP(A:A,Sheet1!A:B,2,0)),"""",VLOOKUP(A:A,Sheet1!A:B,2,0))"
    Next i
    
    columns("O:O").Copy
    columns("O:O").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        
    Range("O:O").NumberFormat = "0"
    
End Sub
