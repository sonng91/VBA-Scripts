Attribute VB_Name = "Addresslookup"
Sub AddressVlookup()
Attribute AddressVlookup.VB_ProcData.VB_Invoke_Func = "R\n14"

    Dim i As Integer
    
    Dim lastRow As Long
        lastRow = Sheets(1).Range("A" & Rows.Count).End(xlUp).Row
        
    Dim i2 As Integer
    
    Dim lastRow2 As Long
        lastRow2 = Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
        
        
    
'First worksheet------------------------------------------------
    Sheets(1).Select
    columns("F").Insert
    
    
    For i = 1 To lastRow
        
        Range("F" & i).Formula = "=LEFT(E:E,8)"
    
    Next i
    
        columns("F:F").Copy
        columns("F:F").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
'2nd Worksheet-------------------------------------------
    
    Sheets("Sheet1").Activate
    
    
    
        columns("B").Insert
    
    For i2 = 1 To lastRow2
        Range("B" & i2).Formula = "=LEFT(A:A,8)"
    
    Next i2
    
        columns("B:B").Copy
        columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
'Back to first worksheet-----------------------------------
    
    Sheets(1).Activate
    
        For i = 2 To lastRow
            Range("N" & i).Formula = "=IF(ISNA(VLOOKUP(F:F,Sheet1!B:C,2,0)),"""",VLOOKUP(F:F,Sheet1!B:C,2,0))"
        Next i
        
        columns("N:N").Copy
        columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    Range("N:N").NumberFormat = "0"
    
    GoTo Carrier

'Carrier Lookup-----------------------------------------------
Carrier:

i = 1

'Range("L" & i).Value = "" And

Do While i < Int(lastRow) 'While the carrier cell is blank, run the code below
    For i = 2 To Int(lastRow)
    x = Range("N" & i).Value


    'CARRIER: UPS
        If InStr(1, x, "1z") > 0 Or InStr(1, x, "1Z") > 0 Then
            If Len(x) = 18 Then

                Range("M" & i).Value = "UPS"
            Else
                Range("M" & i).Value = "Invalid"
            End If
        
        ElseIf Len(x) = 9 Then
            Range("M" & i).Value = "UPS"
        
    'CARRIER: FEDEX
        ElseIf IsNumeric(Range("N" & i)) = "True" Then
            If Len(x) = 15 Or Len(x) = 12 Or Len(x) = 20 Then
                Range("M" & i).Value = "Fedex"
        
        
        
    'CARRIER: FEDEX 14CHAR
            ElseIf Len(x) = 14 Or 13 Then

                Range("L" & i).NumberFormat = "000000000000000"
                Range("M" & i).Value = "Fedex"
    
    'CARRIER: USPS 26CHAR
            ElseIf Len(x) = 26 Then
                Range("M" & i).Value = "USPS"

            
    'CARRIER: FEDEX OR USPS
        
           ElseIf Len(x) = 22 Then
            Range("M" & i).Value = "Fedex or USPS"
           
           End If
        
        
    'CARRIER: CRANE
        ElseIf InStr(1, x, "DSEA") > 0 Or InStr(1, x, "dsea") > 0 Then
            Range("M" & i).Value = "Crane"
            
        
    'Carrier: EMPTY TRACKING
        ElseIf Range("N" & i).Value = "" Then
            Range("M" & i).Value = ""
            
    'CARRIER: Invalid
        Else
            Range("M" & i).Value = "Invalid"
            
            
        End If
        
    Next i
    
    Loop

GoTo hiding

'HidingRows------------------------------------------------------

hiding:

    Range("A1:N1").AutoFilter Field:=13, Criteria1:=Array( _
    "Crane", "Fedex", "UPS", "USPS", "Fedex or USPS"), Operator:=xlFilterValues
    
GoTo columns

'hiding columns--------------------------------------------------
columns:


    columns("C:K").Hidden = True
    

End Sub
