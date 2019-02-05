Sub StockAnalysis()

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

''Define WS Variable
Dim WS As Integer

''Loop Worksheets
For WS = 1 To ActiveWorkbook.Worksheets.Count
    Worksheets(WS).Select
    
    ''Generate Unique Ticker Values
    Dim TickerCount As Double
        TickerCount = 1
    Dim a As Long
        
    ''Define The Variables
    Dim MinValue As Double
    Dim MaxValue As Double
    Dim TotalVolume As Double
        TotalVolume = 0
    Dim IncreaseTicker As String
    Dim IncreaseValue As Double
        IncreaseValue = -10000
    Dim DecreaseTicker As String
    Dim DecreaseValue As Double
        DecreaseValue = 10000
    Dim VolumeTicker As String
    Dim VolumeValue As Double
        VolumeValue = 0
    
    'Create Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    
    'Populate the first row
    Cells(2, 9).Value = Cells(2, 1).Value
    TickerCount = TickerCount + 1
    MinValue = Cells(2, 3).Value
    TotalVolume = TotalVolume + Cells(2, 7).Value
    
    
    'Loop throught the remaining rows
    For a = 3 To Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    
    'Check to see of the ticker is different
        If Cells(a, 1).Value <> Cells(a - 1, 1) Then
            'Add to the Ticker Count
            TickerCount = TickerCount + 1
            
            'Add the ticker symbol to the list
            Cells(TickerCount, 9).Value = Cells(a, 1).Value
            
            'Look back to the previous row to get the max value for the last symbol
            MaxValue = Cells(a - 1, 6)
            
            'Populate the difference in J
            Cells(TickerCount - 1, 10).Value = MaxValue - MinValue
            
            'Add Conditional Formatting
            If Cells(TickerCount - 1, 10).Value > 0 Then
                Cells(TickerCount - 1, 10).Interior.Color = vbGreen
            Else
                Cells(TickerCount - 1, 10).Interior.Color = vbRed
            End If
            
            'Populate the percentage differnece in K
            If MinValue = 0 Then
                Cells(TickerCount - 1, 11).Value = "0"
            Else
                Cells(TickerCount - 1, 11).Value = (MaxValue - MinValue) / MinValue
            End If
            
            'Test for greatest increase/decrease
            If Cells(TickerCount - 1, 11).Value > IncreaseValue Then
                IncreaseValue = Cells(TickerCount - 1, 11).Value
                IncreaseTicker = Cells(TickerCount - 1, 9).Value
            End If
            
            If Cells(TickerCount - 1, 11).Value < DecreaseValue Then
                DecreaseValue = Cells(TickerCount - 1, 11).Value
                DecreaseTicker = Cells(TickerCount - 1, 9).Value
            End If
            
            'Populate Total Volume in L
            Cells(TickerCount - 1, 12).Value = TotalVolume
            
            'Test for greatest volume
            If Cells(TickerCount - 1, 12).Value > VolumeValue Then
                VolumeValue = Cells(TickerCount - 1, 12).Value
                VolumeTicker = Cells(TickerCount - 1, 9).Value
            Else
            End If
            
            'Reset the Min
            MinValue = Cells(a, 3)
            
            'Reset Total Volume
            TotalVolume = Cells(a, 7)
            
            
        Else
            'Add to total Volume
            TotalVolume = TotalVolume + Cells(a, 7)
            
        End If
    
    Next a
    
    'Populate Increase/Decrease/Volume Fields
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 16).Value = IncreaseTicker
    Cells(2, 17).Value = IncreaseValue
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 16).Value = DecreaseTicker
    Cells(3, 17).Value = DecreaseValue
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 16).Value = VolumeTicker
    Cells(4, 17).Value = VolumeValue

Next WS

End Sub
