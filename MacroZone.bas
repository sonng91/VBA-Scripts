Attribute VB_Name = "Module1"
Sub ZoneRates()

Dim OrigZip As String
Dim CusZip As String

Dim i As Double
Dim j As Integer
Dim y As Integer

Dim lastRow As Double
    
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
'Have zone rates in sheet(2)
Dim lastCol As Long
    lastCol = Sheets(2).Cells(3, Columns.Count).End(xlToLeft).Column
    
    
i = 2

For i = 2 To lastRow

'Customer zipcodes in Column F first sheet
'Vendor Zipcode (only the first 3 digits) in Columne E first sheet

'Fixing Zips that begins with leading 0s
        FixCZip = Sheets(1).Range("B" & i)
            If Len(FixCZip) < 5 Then
                FixCZip = Format(FixCZip, "00000")
            End If
        
        FixOZip = Sheets(1).Range("A" & i)
            If Len(FixOZip) < 5 Then
                FixOZip = Format(FixOZip, "00000")
            End If
    
CusZip = Left(FixCZip, 3)
OrigZip = Left(FixOZip, 3)

    For Each Cell In Cells(1, lastCol)
        j = 2
        
        'Put Zipcode Zone in Sheet(2)."B1"
        While j < lastCol + 1
            If OrigZip = Sheets(2).Cells(1, j).Value Then
                
                'Dim LastZR As Integer
                
                ' Found Column Zip
                Dim ColZR As Range
                Set ColZR = Sheets(2).Cells(1, j).Offset(0, -1)
                
                y = 1
                zone1 = False
                Do
                
                    
                    'Dimming variables
                    Dim DesZip As String
                    DesZip = ColZR.Offset(y, 0).Value
                        If Len(DesZip) >= 3 Then
                        
                            Dim beg As String
                            Dim fin As String
                                                        
                            beg = Left(DesZip, 3)
                            fin = Right(DesZip, 3)
                            diff = fin - beg + 1
                            counter = 0
                                Do
                                    'While counter < diff
                                        varx = beg + counter
                                        If Len(varx) < 3 Then
                                            vary = Format(varx, "000")
                                            Else: vary = varx
                                        End If
                                        
                                        If CusZip = vary Then
                                            
                                            
                                            'Writes zone in column N
                                            zone = ColZR.Offset(y, 1).Value
                                            Sheets(1).Range("D" & i).Value = zone
                                            
                                            'MsgBox "The Zone is " & ColZR.Offset(y, 1).Value
                                            
                                            zone1 = True
                                            
                                            Else: counter = counter + 1
                                        
                                        End If
                                        
                                    'Wend
                                Loop Until counter >= diff Or zone1 = True
                            End If
                        y = y + 1
                    Loop Until zone1 = True Or y > 200
                    
                
                      
            End If
            
            j = j + 2
        
        Wend
    Next Cell

Next i



End Sub
