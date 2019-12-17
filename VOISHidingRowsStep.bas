Attribute VB_Name = "VOISHidingRowsStep"
Sub VOIShiderows()
'Module7
'hiding empty rows


    Range("A1:M1").AutoFilter Field:=12, Criteria1:=Array( _
    "Crane", "Fedex", "UPS", "USPS", "Fedex or USPS"), Operator:=xlFilterValues
    
    
End Sub
