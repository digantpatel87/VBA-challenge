
Sub Main()

For Each ws In Worksheets
    Dim MaxRow As Double
    
    MaxRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    'Populate distinct Tickers
    Dim NextTickerCell As Integer
    NextTickerCell = 2
    Dim RunningTicker As String
    Dim RunningDistTicker As String
    Dim RunningOpen As Double
    Dim RunningClose As Double
    Dim RunningVolumn As Integer
    Dim MaxnumberOfGivenTicker As Integer
    Dim PreviousMaxnumberOfGivenTicker As Double
    Dim RunningTotal As Double
    
    PreviousMaxnumberOfGivenTicker = 0
    
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percentage Change"
    ws.Cells(1, "L").Value = "Total Sock Volume"
    
    ws.Cells(2, "O").Value = "Greatest % increase"
    ws.Cells(3, "O").Value = "Greatest % decrease"
    ws.Cells(4, "O").Value = "Greatest total volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    
    'Set initial Start RunningOpen
    RunningOpen = ws.Cells(NextTickerCell, "C").Value
    
    'Loop for each row
    For i = 2 To MaxRow
        RunningTicker = ws.Cells(i, "A")
        
        'Check if ticker exists in I column
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Set new ticker in I column
            ws.Cells(NextTickerCell, "I").Value = RunningTicker
            
            'as total
            RunningTotal = RunningTotal + ws.Cells(i, "G")
           
            'get max records of given ticker
            MaxnumberOfGivenTicker = Application.WorksheetFunction.CountIf(Range("A2:A" & MaxRow), RunningTicker)
            
            'Add max given ticker records to running number to use it to get closing number
            'RunningClose = Cells(1 + MaxnumberOfGivenTicker + PreviousMaxnumberOfGivenTicker, "F").Value
            RunningClose = ws.Cells(i, "F").Value
            
            'Get Difference
            ws.Cells(NextTickerCell, "J").Value = RunningClose - RunningOpen
            
           If RunningOpen <> 0 Then
                'Calculate percentage change
                ws.Cells(NextTickerCell, "K").Value = (RunningClose - RunningOpen) / RunningOpen
                ws.Cells(NextTickerCell, "K").NumberFormat = "0.00%"
            Else
                ws.Cells(NextTickerCell, "K").Value = 0
                ws.Cells(NextTickerCell, "K").NumberFormat = "0.00%"
            End If
            
            'Get the sum of given Ticker using Sumif
             ws.Cells(NextTickerCell, "L").Value = RunningTotal
            
            'Set this to be used for getting close number
            PreviousMaxnumberOfGivenTicker = MaxnumberOfGivenTicker + PreviousMaxnumberOfGivenTicker
            'Set this to determine next cell for I column
            NextTickerCell = NextTickerCell + 1
            
            'reset Total
            RunningTotal = 0
             'Get Open of next ticker
            RunningOpen = ws.Cells(i + 1, "C").Value
          Else
            RunningTotal = RunningTotal + ws.Cells(i, "G")
                                          
        End If
        
              
    Next i
          
    'Conditional formatting
    'clear any existing conditional formatting
    ws.Range("J2", ws.Range("J2").End(xlDown)).FormatConditions.Delete

    'Conditional Formating for Positive value in Green color
    With ws.Range("J2", ws.Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        .Interior.Color = vbGreen
    End With
    
    'Conditional Formating for Negative value in red color
    With ws.Range("J2", ws.Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlLess, "=0")
        .Interior.Color = vbRed
    End With
    
    
    'Bonus
    Dim BonusMaxRow As Long
    Dim BonusRunningTicker As String
    Dim BonusRunningPercentage As Double
    Dim BonusRunningTotalStockVolumn As Double
    Dim BonusTotalStockVolumn As Double
           
        
        
    BonusMaxRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
      
    ws.Cells(2, "Q").Value = 0
    ws.Cells(3, "Q").Value = 0
    ws.Cells(4, "Q").Value = 0
    
    ws.Cells(2, "Q").NumberFormat = "0.00%"
    ws.Cells(3, "Q").NumberFormat = "0.00%"
    
    
    For j = 2 To BonusMaxRow
        BonusRunningTicker = ws.Cells(j, "I")
        BonusRunningPercentage = ws.Cells(j, "K")
        BonusRunningTotalStockVolumn = ws.Cells(j, "L")
        
        If BonusRunningPercentage > ws.Cells(2, "Q").Value Then
            ws.Cells(2, "Q").Value = BonusRunningPercentage
            ws.Cells(2, "P").Value = BonusRunningTicker
        End If
        
        If BonusRunningPercentage < ws.Cells(3, "Q").Value Then
            ws.Cells(3, "Q").Value = BonusRunningPercentage
            ws.Cells(3, "P").Value = BonusRunningTicker
        End If
        
        
        BonusTotalStockVolumn = ws.Cells(4, "Q").Value
        
        If BonusRunningTotalStockVolumn > BonusTotalStockVolumn Then
            ws.Cells(4, "Q").Value = BonusRunningTotalStockVolumn
            ws.Cells(4, "P").Value = BonusRunningTicker
        End If
        
    
    Next j

Next ws
    
End Sub



