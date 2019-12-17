Attribute VB_Name = "VOISCarrierStep"
Sub vendorcarrier3()
'


Dim i As Integer
Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
Dim x As String 'Determines Cell position
Dim y As String 'for UPS case sensitive search



i = 1

'Range("L" & i).Value = "" And

Do While i < Int(lastRow) 'While the carrier cell is blank, run the code below
    For i = 2 To Int(lastRow)
    x = Range("M" & i).Value


    'CARRIER: UPS
        If InStr(1, x, "1z") > 0 Or InStr(1, x, "1Z") > 0 Then
            If Len(x) = 18 Then

                Range("L" & i).Value = "UPS"
            Else
                Range("L" & i).Value = "Invalid"
            End If
        
        ElseIf Len(x) = 9 Then
            Range("L" & i).Value = "UPS"
        
    'CARRIER: FEDEX
        ElseIf IsNumeric(Range("M" & i)) = "True" Then
            If Len(x) = 15 Or Len(x) = 12 Or Len(x) = 20 Then
                Range("L" & i).Value = "Fedex"
        
        
    'CARRIER: USPS 26CHAR
            ElseIf Len(x) = 26 Then
                Range("L" & i).Value = "USPS"
                
                
    'CARRIER: FEDEX 14CHAR
            ElseIf Len(x) = 14 Or Len(x) = 13 Then

                Range("M" & i).NumberFormat = "000000000000000"
                Range("L" & i).Value = "Fedex"

            
    'CARRIER: FEDEX OR USPS
        
           ElseIf Len(x) = 22 Then
            Range("L" & i).Value = "Fedex or USPS"
           
           End If
        
        
    'CARRIER: CRANE
        ElseIf InStr(1, x, "DSEA") > 0 Or InStr(1, x, "dsea") > 0 Then
            Range("L" & i).Value = "Crane"
            
        
    'Carrier: EMPTY TRACKING
        ElseIf Range("O" & i).Value = "" Then
            Range("L" & i).Value = ""
            
    'CARRIER: Invalid
        Else
            Range("L" & i).Value = "Invalid"
            
            
        End If
        
    Next i
    
    Loop
    
'MsgBox "Ellapsed Time in Hrs:Min:Sec :" & Format(Now() - t, "hh:mm:ss")

End Sub

