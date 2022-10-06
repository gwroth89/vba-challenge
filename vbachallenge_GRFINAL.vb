Sub stock_analysis():
   
   ' Set variables
   Dim i As Long
   Dim ticker As String
   Dim volume As Double
   Dim change As Double
   Dim j As Long
   Dim rowCount As Long
   Dim percentChange As Double
   Dim start As Double
   
   ' Set headers
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Volume"
   Range("P1").Value = "Ticker"
   Range("Q1").Value = "Value"
   Range("O2").Value = "Greatest % Increase"
   Range("O3").Value = "Greatest % Decrease"
   Range("O4").Value = "Greatest Total Volume"
   
   ' Set values
   j = 0
   volume = 0
   change = 0
   start = 2
   rowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
   For i = 2 To rowCount
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            volume = volume + Cells(i, 7).Value
           ' Handle zero total volume
           If volume = 0 Then

               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = 0
               Range("K" & 2 + j).Value = "%" & 0
               Range("L" & 2 + j).Value = 0
           Else
               ' Find First non zero starting value
               If Cells(start, 3) = 0 Then
                   For find_value = start To i
                       If Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                       End If
                    Next find_value
               End If
               'Calculate total change and percent change columns
               change = (Cells(i, 6) - Cells(start, 3))
               percentChange = Round((change / Cells(start, 3) * 100), 2)
               'next ticker
               start = i + 1

               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = Round(change, 2)
               Range("K" & 2 + j).Value = "%" & percentChange
               Range("L" & 2 + j).Value = volume
           End If

            volume = 0
            change = 0
            j = j +1

       Else
       volume = volume + Cells(i, 7).Value
       End If
   Next i
End Sub

'Commented sub below was final attempt at formatting the change column
'Sub formatting()

'Dim rowCount as Integer
'Dim percentChange As Double

'rowCount = Cells(Rows.Count, "A").End(xlUp).Row
'percentChange = Range("J").Value

'From i = 2 to rowCount
  ' Check if ticker change is positive
    'If Range("J").Value > 0 Then

        ' Color the change green
        'Range("J").Interior.ColorIndex = 4

  
        'Else if percentChange < 0 Then

        ' Color the change red
        'Range("J").Interior.ColorIndex = 3

        'End If
    'Next i 
'End Sub