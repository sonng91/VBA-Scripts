Attribute VB_Name = "StatesList"
Sub States()

    Dim c As Object
    Set c = CreateObject("scripting.Dictionary")
   

'Add some (key, value) pairs
c.Add "Alabama", "AL"
c.Add "Alaska", "AK"
c.Add "Arizona", "AZ"
c.Add "Arkansas", "AR"
c.Add "California", "CA"
c.Add "Colorado", "CO"
c.Add "Connecticut", "CT"
c.Add "Delaware", "DE"
c.Add "Florida", "FL"
c.Add "Georgia", "GA"
c.Add "Hawaii", "HI"
c.Add "Idaho", "ID"
c.Add "Illinois", "IL"
c.Add "Indiana", "IN"
c.Add "Iowa", "IA"
c.Add "Kansas", "KS"
c.Add "Kentucky", "KY"
c.Add "Louisiana", "LA"
c.Add "Maine", "ME"
c.Add "Maryland", "MD"
c.Add "Massachusetts", "MA"
c.Add "Michigan", "MI"
c.Add "Minnesota", "MN"
c.Add "Mississippi", "MS"
c.Add "Missouri", "MO"
c.Add "Montana", "MT"
c.Add "Nebraska", "NE"
c.Add "Nevada", "NV"
c.Add "New Hampshire", "NH"
c.Add "New Jersey", "NJ"
c.Add "New Mexico", "NM"
c.Add "New York", "NY"
c.Add "North Carolina", "NC"
c.Add "North Dakota", "ND"
c.Add "Ohio", "OH"
c.Add "Oklahoma", "OK"
c.Add "Oregon", "OR"
c.Add "Pennsylvania", "PA"
c.Add "Rhode Island", "RI"
c.Add "South Carolina", "SC"
c.Add "South Dakota", "SD"
c.Add "Tennessee", "TN"
c.Add "Texas", "TX"
c.Add "Utah", "UT"
c.Add "Vermont", "VT"
c.Add "Virginia", "VA"
c.Add "Washington", "WA"
c.Add "West Virginia", "WV"
c.Add "Wisconsin", "WI"
c.Add "Wyoming", "WY"
c.Add "District of Columbia", "DC"
c.Add "Virgin Islands", "VI"
c.Add "Puerto Rico", "PR"
c.Add "Guam", "GU"

'c.Add "Item", Key

    Dim lastRow As Long

    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    Dim StrCell As String
    
    i = 1

Do While i < Int(lastRow)
Backup:
For i = i To lastRow
StrCell = Range("A" & i).Value
    'If StrCell = "" Then
    '    i = i + 1
    '    GoTo Backup
    If Len(StrCell) = 2 Then
        Range("B" & i).Value = StrCell
    ElseIf c.Exists(StrCell) Then 'if StrCell string exist in the dictionary "c.Add 'Item', Key"
        Range("B" & i).Value = c.Item(StrCell)
        'Range("B" & i).Value = c.Item(StrCell) 'will put in the "Key"
    End If
    If Not c.Exists(StrCell) Then
        Range("B" & i).Value = StrCell
    End If
    
Next i
Loop

End Sub
