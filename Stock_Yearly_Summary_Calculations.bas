Attribute VB_Name = "Calculation_Module"
Sub calc_ticker_changes()

    'Read ticker data orderd by ticker symbol and date.
    'For each ticker symbol calculate the folowing:
    
    'Yearly Change  = Closing Stock value for Last day of year - Opening Stock Value of First day of year
    
    'Percent Change = (Yearly Change / Opening Stock Value of First Day of Year) x 100
    
    'Total Stock Volume = Sum of all Volume for the entire Year
    
    'Present the the three calculations for each ticker in a separate area of the worksheet.
    
    
    'Store Opening value
    Dim openingValue As Double
    
    'Store Volume Sum for Ticker
    Dim tickerVolume As Double
    
    'Store current working record
    Dim currentRec As Integer
    
    'Current Ticker Display row
    Dim tickerSumRow As Integer
        
    'Greatest Ticker Increase
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestIncreaseVal As Double
    
    'Greatest Ticker Decrease
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestDecreaseVal As Double
    
    'Greatest Ticker Volume
    Dim tickerGreatestVolume As String
    Dim tickerGreatestVolumeVal As Double
    
    tickerGreatestIncreaseVal = 0
    tickerGreatestDecreaseVal = 0
    tickerGreatestVolumeVale = 0
    
    currentRec = 2
    
    tickerSumRow = 2
    
    'Create header titles for new ticker summary area
    Range("I1").Value2 = "Ticker"
    Range("J1").Value2 = "Yearly Change"
    Range("K1").Value2 = "Percent Change"
    Range("L1").Value2 = "Total Volume Change"
    
    Columns(11).NumberFormat = "0.00%"
    Columns(12).NumberFormat = "#,##0"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "#,##0"
    
    Dim lastRow As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For currentRec = 2 To lastRow
    
        If Cells(currentRec + 1, 1).Value <> Cells(currentRec, 1).Value Then
        
            tickerVolume = tickerVolume + Cells(currentRec, 7)
        
            Range("I" & tickerSumRow).Value2 = Cells(currentRec, 1).Value
            
            Range("J" & tickerSumRow).Value2 = Cells(currentRec, 6).Value - openingValue
            
            Range("K" & tickerSumRow).Value2 = (Cells(currentRec, 6).Value - openingValue) / openingValue
            
            Range("L" & tickerSumRow).Value2 = tickerVolume
             
            If Range("J" & tickerSumRow).Value2 < 0 Then
              Range("J" & tickerSumRow).Interior.ColorIndex = 3  ' Red
            Else
              Range("J" & tickerSumRow).Interior.ColorIndex = 4   'Green
            End If
            
            
            If Range("K" & tickerSumRow).Value < tickerGreatestDecreaseVal Then
              tickerGreatestDecreaseVal = Range("K" & tickerSumRow).Value
              tickerGreateastDecrease = Cells(currentRec, 1).Value
            ElseIf Range("K" & tickerSumRow).Value > tickerGreatestIncreaseVal Then
              tickerGreatestIncreaseVal = Range("K" & tickerSumRow).Value
              tickerGreateastIncrease = Cells(currentRec, 1).Value
            End If
            
            If Range("L" & tickerSumRow).Value > tickerGreatestVolumeVal Then
              tickerGreatestVolumeVal = Range("L" & tickerSumRow).Value
              tickerGreateastVolume = Cells(currentRec, 1).Value
            End If
            
            
            tickerVolume = 0
            openingValue = Cells(currentRec + 1, 3).Value2
            tickerSumRow = tickerSumRow + 1
                
        Else
        
            tickerVolume = tickerVolume + Cells(currentRec, 7).Value
            
            If currentRec = 2 Then
              openingValue = Cells(currentRec, 3).Value2
            End If
            
        End If
    
    Next currentRec
    
    Range("P2").Value = tickerGreatestIncrease
    Range("Q2").Value = tickerGreatestIncreaseVal
    
    Range("P3").Value = tickerGreatestDecrease
    Range("Q3").Value = tickerGreatestDecreaseVal

    Range("P4").Value = tickerGreatestVolume
    Range("Q4").Value = tickerGreatestVolumeVal
    
    'Autofit Column Widths to Data
    Columns("I:Q").AutoFit

End Sub
