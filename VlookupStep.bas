Attribute VB_Name = "VlookupStep"
Sub Vlookup()
Attribute Vlookup.VB_ProcData.VB_Invoke_Func = " \n14"
'
'Module5
'
'vlookup macro for dropship tableau
'
'
'
'

'Dim t As Date
'set a variable equal to the starting time
't = Now()


Dim lastRow As Long
    lastRow = Range("B" & Rows.Count).End(xlUp).Row

    
    Range("H1").Value = "Status"
    
    Range("G2").Formula = "=IF(ISNA(VLOOKUP(C2,J:K,2,0)),"""",VLOOKUP(C2,J:K,2,0))"
    Range("G2").AutoFill Destination:=Range("G2:G" & lastRow)
    
    Range("H2").Formula = "=IF(ISNA(VLOOKUP(C2,J:L,3,0)),"""",VLOOKUP(C2,J:L,3,0))"
    Range("H2").AutoFill Destination:=Range("H2:H" & lastRow)
    
    
    columns("G:G").Copy
    columns("G:G").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    columns("H:H").Copy
    columns("H:H").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("G:G").NumberFormat = "0"
    
'MsgBox "Ellapsed Time in Hrs:Min:Sec :" & Format(Now() - t, "hh:mm:ss")

End Sub
