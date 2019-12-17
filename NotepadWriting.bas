Attribute VB_Name = "NotepadWriting"
Sub notepadwriteOrig()

Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
Dim lastCol As Long
    lastCol = Cells(1, columns.Count).End(xlToLeft).Column
    
Dim c As Object
Set c = CreateObject("scripting.Dictionary")
    
Dim i As Integer
Dim j As Integer
Dim final As String

Dim title As String

title = Range("B2").Value


'Copying each row and storing in memory

i = 2
While i < lastRow + 1
    
    For Each Cell In Cells(i, lastCol)
    
        j = 1
        While j < lastCol + 1
            
            final = final + CStr(Cells(i, j).Value)
            j = j + 1
        Wend
        
        'adding to dictionary
        c.Add i, final
        'MsgBox "Key: " & i & " " & c.Item(i)
        final = ""
        'MsgBox c.Count
        i = i + 1
        
    Next Cell
    
    'CreateObject("scripting.filesystemobject").createtextfile("C:\Users\snguyen\Downloads\test.txt").write final
    
Wend

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = fso.CreateTextFile("C:\Users\snguyen\Downloads\QVC_" & title & ".txt") 'Modify as needed
oFile.Write PrintDictionary(c)           'modify as needed
Set fso = Nothing: Set oFile = Nothing

End Sub

Function PrintDictionary(c As Object) As String
    
    Dim k As Variant
    Dim i As Long
    Dim fullText As String
    
    For Each k In c.Keys()
        fullText = fullText & c(k) & vbCrLf
    Next
    
    PrintDictionary = fullText


End Function

