Attribute VB_Name = "Module16"
Sub Newvlookup()
'Module 16

'creating new worksheet
'creates a list of unique values
'
'

Dim t As Date
'set a variable equal to the starting time
t = Now()

Dim i As Double
Dim lastRow As Long
    lastRow = Range("C" & Rows.Count).End(xlUp).Row
Dim uniqueLast As Long
    uniqueLast = Range("J" & Rows.Count).End(xlUp).Row

Dim found As Boolean
Dim x As Integer
uniquepos = 2

x = 1
found = False
    
'Range("B1").Value = "trackings"

'--------------------------------------'
    
For i = 2 To lastRow
    If IsEmpty(Range("A" & i).Value) Then GoTo nexti
    
    If x < uniqueLast Then
        Do
            If Range("C" & i) = Range("J" & x) Then
    
                Range("J" & x).Offset(, 1).Copy Range("G" & i)
                found = True
                
    
            End If
        'goes to next cell to see if it's a duplicate in new list
        x = x + 1
        Loop Until x = Int(uniqueLast) Or found = True
    
    End If



    If found = True Then GoTo nexti
    
        'temp = Range("A" & i).Value
        
    'Range("A" & i).Copy Range("B" & uniquepos)
        'Range("D" & uniquepos).Value = temp
        
        
    'decides where to put in the new unique value
        'uniquepos = uniquepos + 1
        
        
    'row count for new list of uniques
        'uniqueLast = uniqueLast + 1
    
nexti:
    'restarts the search in the new list for any duplicates
        x = 1
        
    'resets Boolean Test
        found = False
    
    Next i
    

MsgBox "Elapsed Time in Hrs:Min:Sec :" & Format(Now() - t, "hh:mm:ss")

End Sub

