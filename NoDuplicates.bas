Attribute VB_Name = "NoDuplicates"
Sub UniqueValues()
Attribute UniqueValues.VB_Description = "Unique Values"
Attribute UniqueValues.VB_ProcData.VB_Invoke_Func = "U\n14"
'Module 14

'creating new worksheet
'creates a list of unique values
'Shortcut: CTRL+SHIFT+U
'
Dim i As Double
Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    

Dim found As Boolean
Dim x As Integer
uniquepos = 2
uniqueLast = 2
x = 1
found = False
    
Range("B1").Value = "Uniques"

'--------------------------------------'
    
For i = 2 To lastRow
    If IsEmpty(Range("A" & i).Value) Then GoTo nexti
    
    If x < uniqueLast Then
        Do
            If Range("A" & i) = Range("B" & x) Then
    
                'Searching through the list for same value
                found = True
                
    
            End If
        'goes to next cell to see if it's a duplicate in new list
        x = x + 1
        Loop Until x = uniqueLast Or found = True
    
    End If

    If found = True Then GoTo nexti
    
        temp = Range("A" & i).Value
        
        Range("B" & uniquepos).Value = temp
        
        
    'decides where to put in the new unique value
        uniquepos = uniquepos + 1
        
        
    'row count for new list of uniques
        uniqueLast = uniqueLast + 1
    
nexti:
    'restarts the search in the new list for any duplicates
        x = 1
        
    'resets Boolean Test
        found = False
    
    Next i
    
    
End Sub
