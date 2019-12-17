Attribute VB_Name = "VoisVendorMacro"
Sub VOISmacro()
Attribute VOISmacro.VB_ProcData.VB_Invoke_Func = "V\n14"
'Module20


Dim t As Date
'set a variable equal to the starting time
t = Now()

    Sheets(1).Activate
    Sheets(2).name = "Sheet1"
    If Range("L1").Value = "Staged Count" Then
        columns("I:I").Delete
        columns("K:K").Delete
    End If
    
    If Not Range("A1").Value = "Order Id" Or Not Range("M1").Value = "Tracking Number" Then
        MsgBox "This isn't the right file."
        GoTo finished
    End If
    
    'Call trimmingTrackings 'trimmingStuff
    
    Call VOISvlookupStep 'Voisvlookup
    
    Call vendorcarrier3 'VoisCarrierStep
    
    Call VOIShiderows 'HidingRowsStep
    
    'Call States 'StatesList
    
    Call HideColumns   'Hide Columns
    
MsgBox "Ellapsed Time in Hrs:Min:Sec :" & Format(Now() - t, "hh:mm:ss")

finished:
End Sub
