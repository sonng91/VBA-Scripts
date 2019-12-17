Attribute VB_Name = "Voisvlookup"
Sub VOISvlookupStep()
'Module13
'
'vlookup macro for dropship tableau
'

Dim i As Integer

Dim lastRow As Long
    lastRow = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
    


'Vlookup-------------------------------------------------------------------------------------------------

Sheets(1).Select

    For i = 2 To lastRow
        Range("M" & i).Formula = "=IF(ISNA(VLOOKUP(A:A,Sheet1!A:B,2,0)),"""",VLOOKUP(A:A,Sheet1!A:B,2,0))"
    Next i
    
    columns("M:M").Copy
    columns("M:M").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        
    Range("M:M").NumberFormat = "0"

End Sub
