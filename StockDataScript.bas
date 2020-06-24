Attribute VB_Name = "Module1"
Sub StockData()
    
    On Error Resume Next
      
    'declare variables
    Dim Ticker As String
    Dim Ticker_StartYV As Double 'Start Year Value
    Dim Ticker_EndYV As Double 'End Year Value
    Dim Ticker_Percent_Change As Double
    Dim Ticker_Total_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim Ticker_Count As Integer
    Dim MaxRow As Integer

    For Each ws In Worksheets
                
        'initialize variables
        Ticker_StartYV = 0
        Ticker_EndYV = 0
        Ticker_Percent_Change = 0
        Ticker_Total_Volume = 0
        Summary_Table_Row = 2
        Ticker_Count = 0
        
        'find last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'find last column
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        'set summary column headers
        ws.Cells(1, LastColumn + 2) = "Ticker"
        ws.Cells(1, LastColumn + 3) = "Yearly Change"
        ws.Cells(1, LastColumn + 4) = "Percent Change"
        ws.Cells(1, LastColumn + 5) = "Total Stock Volume"
    
        'Loop through all rows
        For i = 2 To LastRow
    
            'If the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'update variables
                Ticker = ws.Cells(i, 1).Value
                Ticker_EndYV = ws.Cells(i, 6).Value
                Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
                Ticker_StartYV = ws.Cells(i - Ticker_Count, 3)
            
                'print variables
                ws.Cells(Summary_Table_Row, LastColumn + 2).Value = Ticker
                ws.Cells(Summary_Table_Row, LastColumn + 3).Value = Ticker_EndYV - Ticker_StartYV
                ws.Cells(Summary_Table_Row, LastColumn + 4).Value = (Ticker_EndYV - Ticker_StartYV) / Ticker_StartYV
                ws.Cells(Summary_Table_Row, LastColumn + 5).Value = Ticker_Total_Volume
                
                'format fill color % change column
                If ws.Cells(Summary_Table_Row, LastColumn + 3).Value > 0 Then
                    ws.Cells(Summary_Table_Row, LastColumn + 3).Interior.Color = VBA.ColorConstants.vbGreen
                ElseIf ws.Cells(Summary_Table_Row, LastColumn + 3).Value < 0 Then
                    ws.Cells(Summary_Table_Row, LastColumn + 3).Interior.Color = VBA.ColorConstants.vbRed
                End If
                
                'format percent change column as a percentage
                ws.Cells(Summary_Table_Row, LastColumn + 4).NumberFormat = "0.00%"
            
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                  
                'Reset containers
                Ticker_Total_Volume = 0
                Ticker_Count = 0
        
            'the ticker does not change
            Else
        
                'Add to the Volume sum
                Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
                
                'increment ticker counter
                Ticker_Count = Ticker_Count + 1
        
            End If
    
        Next i
        
        'set column and row headers
        ws.Cells(2, LastColumn + 8).Value = "Greatest % Increase"
        ws.Cells(3, LastColumn + 8).Value = "Greatest % Decrease"
        ws.Cells(4, LastColumn + 8).Value = "Greatest Total Volume"
        ws.Cells(1, LastColumn + 9).Value = "Ticker"
        ws.Cells(1, LastColumn + 10).Value = "Value"
        
        'find last row in the new summary columns
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'declare variables
        Dim GreatestIncrease As Double
        Dim GreatestIncTicker As String
        Dim GreatestDecrease As Double
        Dim GreatestDecTicker As String
        Dim GreatestVolume As Double
        Dim GreatestVolTicker As String
        
        'set variables
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        
        For i = 2 To LastRow
            'Check for Greatest Increase
            If ws.Cells(i, 11) > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, LastColumn + 4)
                GreatestIncTicker = ws.Cells(i, LastColumn + 2)
            End If
            'Check for Greatest Decrease
            If ws.Cells(i, 11) < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, LastColumn + 4)
                GreatestDecTicker = ws.Cells(i, LastColumn + 2)
            End If
            'Check for Greatest Volume
            If ws.Cells(i, 12) > GreatestVolume Then
                GreatestVolume = ws.Cells(i, LastColumn + 5)
                GreatestVolTicker = ws.Cells(i, LastColumn + 2)
            End If
        Next i
        
        'Print variables
        ws.Cells(2, LastColumn + 9) = GreatestIncTicker
        ws.Cells(3, LastColumn + 9) = GreatestDecTicker
        ws.Cells(4, LastColumn + 9) = GreatestVolTicker
        ws.Cells(2, LastColumn + 10) = GreatestIncrease
        ws.Cells(3, LastColumn + 10) = GreatestDecrease
        ws.Cells(4, LastColumn + 10) = GreatestVolume
        
        'this section was my first pass at identifying the greatest changes
        'works without using a loop, but was way more time consuming to develop
        
        'call the worksheet function "MAX" to find the max value for % change
        'ws.Cells(2, LastColumn + 10).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, LastColumn + 4), ws.Cells(LastRow, LastColumn + 4)))
        'find the row that holds the MAX value
        'MaxRow = Application.WorksheetFunction.Match(ws.Cells(2, LastColumn + 10).Value, ws.Range(ws.Cells(1, LastColumn + 4), ws.Cells(LastRow, LastColumn + 4)), 0)
        'identify the ticker associated with the MAX value
        'ws.Cells(2, LastColumn + 9).Value = ws.Cells(MaxRow, LastColumn + 2).Value
        'call the worksheet function "MIN" to find the min value for % change
        'ws.Cells(3, LastColumn + 10).Value = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, LastColumn + 4), ws.Cells(LastRow, LastColumn + 4)))
        'find the row that holds the MIN value
        'MaxRow = Application.WorksheetFunction.Match(ws.Cells(3, LastColumn + 10).Value, ws.Range(ws.Cells(1, LastColumn + 4), ws.Cells(LastRow, LastColumn + 4)), 0)
        'identify the ticker associated with the MIN value
        'ws.Cells(3, LastColumn + 9).Value = ws.Cells(MaxRow, LastColumn + 2).Value
        'call the worksheet function "MAX" to find the max value for volume
        'ws.Cells(4, LastColumn + 10).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, LastColumn + 5), ws.Cells(LastRow, LastColumn + 5)))
        'find the row that holds the MAX volume
        'MaxRow = Application.WorksheetFunction.Match(ws.Cells(4, LastColumn + 10).Value, ws.Range(ws.Cells(1, LastColumn + 5), ws.Cells(LastRow, LastColumn + 5)), 0)
        'identify the ticker associated with the MAX volume
        'ws.Cells(4, LastColumn + 9).Value = ws.Cells(MaxRow, LastColumn + 2).Value
        
        'format % columns
        ws.Cells(2, LastColumn + 10).NumberFormat = "0.00%"
        ws.Cells(3, LastColumn + 10).NumberFormat = "0.00%"
        
    Next ws

End Sub
