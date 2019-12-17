Attribute VB_Name = "Module1"
Sub testDictionary()

    Dim c As Object
    Set c = CreateObject("Scripting.Dictionary")
    
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

    
    Dim Cell As Variant
    
    Dim i As Integer
    
    
    i = 1
    For i = 1 To 50
        Cell = Range("A" & i)
        If c.Exists(Cell) Then
            
            MsgBox c.Items(Cell) & " exists."
            
        End If
    Next i


End Sub
