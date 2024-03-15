Attribute VB_Name = "Module1"
Sub Stockdata_Analysis()

'Loop through all worksheets in the workbook
For Each ws In ThisWorkbook.Worksheets

'Declare variables to store relevant information
    Dim Ticker As String
    Dim YealyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim TickerRow As Long: TickerRow = 2
    Dim SummaryRow As Long
    Dim OpeningPice As Double
    Dim ClosingPrice As Double
    
'Insert Column Headings into specific cells
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
   OpeningPrice = ws.Cells(2, 3).Value
    
'Find the last row of the data
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    
 'Loop through all rows (loops will start from row 2 because row 1 contains headers)
    'Calculation for YearlyChange And TotalVolume
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(TickerRow, 9).Value = Ticker
                       
        
            ClosingPrice = ws.Cells(i, 6).Value
            YearlyChange = ClosingPrice - OpeningPrice
            ws.Cells(TickerRow, 10).Value = YearlyChange
            
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Cells(TickerRow, 12).Value = TotalVolume
            
            'Calculation for Percent Change
            If OpeningPrice <> 0 Then
                  PercentChange = (YearlyChange / OpeningPrice)
            Else
                  PercentChange = 0
                      
            End If
            ws.Cells(TickerRow, 11).Value = PercentChange
               
            ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
            
            'use conditional formating to highlight positive vs negative Yearly change
            
             If ws.Cells(TickerRow, 10).Value > 0 Then
                ws.Cells(TickerRow, 10).Interior.Color = RGB(0, 255, 0) 'Green
    
             ElseIf ws.Cells(TickerRow, 10).Value < 0 Then
                 ws.Cells(TickerRow, 10).Interior.Color = RGB(255, 0, 0) 'Red
    
           End If
           
            'Reset TotalVolume for the next ticker
            TotalVolume = 0
            'Updage TickerRow for next ticker
            TickerRow = TickerRow + 1
              
    
            OpeningPrice = ws.Cells(i + 1, 3).Value
        Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
        End If
    
    Next i
    'Add functionality to Stockdata to show Greatest % increase, decrease & total volume
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'create calculation for max
    ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    
   Row_MaxIncrease = WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("K2:K" & LastRow), 0)
   Row_MaxDecrease = WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("K2:K" & LastRow), 0)
   Row_MaxVolume = WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L2:L" & LastRow), 0)
   
   ws.Cells(2, 16).Value = ws.Cells(Row_MaxIncrease + 1, 9).Value
    ws.Cells(3, 16).Value = ws.Cells(Row_MaxDecrease + 1, 9).Value
    ws.Cells(4, 16).Value = ws.Cells(Row_MaxVolume + 1, 9).Value
    
Next ws
    
 
       
End Sub
