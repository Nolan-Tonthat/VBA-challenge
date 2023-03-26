Attribute VB_Name = "Module1"
Sub StockFormat()
For Each ws In Worksheets
    Dim WorksheetName As String
    
    'current row
    Dim CurrRow As Long
    
    'start row of current Tick
    Dim TickStart As Long
    
    '# of current Ticks
    Dim TickCount As Long
    
    'Last Row of Original Tick Data
    Dim LastRow As Long
    
    
    '% change
    Dim PerChange As Double
    

    WorksheetName = ws.Name

    'column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

        
        'Set Ticker Counter to first row
        TickCount = 2
        
        'Set start row to 2
        TickStart = 2
        
        'Find the last data cell in column A
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows
            For CurrRow = 2 To LastRow
            
                'Create new tick name in summary column if tick name changes
                If ws.Cells(CurrRow + 1, 1).Value <> ws.Cells(CurrRow, 1).Value Then
                ws.Cells(TickCount, 9).Value = ws.Cells(CurrRow, 1).Value
                
                'Compares current row close value with initial open value.  Ends calculation with close value of last row of current tick to get Yearly Change value.
                ws.Cells(TickCount, 10).Value = ws.Cells(CurrRow, 6).Value - ws.Cells(TickStart, 3).Value
                
                    'Set background to Green (4) if change is positive. Set background to Red (3) if change is negative.
                    If ws.Cells(TickCount, 10).Value > 0 Then
                        ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                    End If
                    
                    'Calculate and write percent change in column K
                    If ws.Cells(TickStart, 3).Value <> 0 Then
                        PerChange = ((ws.Cells(CurrRow, 6).Value - ws.Cells(TickStart, 3).Value) / ws.Cells(TickStart, 3).Value)
                        ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent") 'Formatting column K to add % sign
                    Else
                        ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    End If
                    
                'Calculate and write total volume in column L
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(TickStart, 7), ws.Cells(CurrRow, 7)))
                
                'Increase TickCount by 1
                TickCount = TickCount + 1
                
                'Set new start row of new
                TickStart = CurrRow + 1
                
                End If
            
            Next CurrRow
            
        
    Next ws

End Sub

Sub SummaryWholeDataSet()
    
    Dim WorksheetName As String
    
    'current row
    Dim CurrRow As Long
    
    'Last Row of Concatenated Tick Data
    Dim SummRows As Long
    
    'Highest increase
    Dim HighestIncr As Double
    
    'Highest decrease
    Dim HighestDecr As Double
    
    'Highest volume
    Dim HighestVol As Double
  
    
        
        
    For Each ws In Worksheets
        'Find last non-blank cell in column I
        SummRows = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        WorksheetName = ws.Name
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        HighestVol = ws.Cells(2, 12).Value
        HighestIncr = ws.Cells(2, 11).Value
        HighestDecr = ws.Cells(2, 11).Value
        
            'Loop for summary
            For CurrRow = 2 To SummRows
            
                
                'Greatest increase.  Compares all increases one by one.  If new increase is greater, populate cell with new increase.
                If ws.Cells(CurrRow, 11).Value > HighestIncr Then
                HighestIncr = ws.Cells(CurrRow, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(CurrRow, 9).Value
                
                ElseIf ws.Cells(CurrRow, 11).Value < HighestDecr Then
                
                
                'Greatest decrease.  Compares all decreases one by one.  If new decrease is greater, populate cell with new decrease.
                
                HighestDecr = ws.Cells(CurrRow, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(CurrRow, 9).Value
                
                ElseIf ws.Cells(CurrRow, 12).Value > HighestVol Then
                
                'Greatest total volume.  Compares all total volumes one by one.  If new total volume is greater, populate cell with new total volume.

                HighestVol = ws.Cells(CurrRow, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(CurrRow, 9).Value
                
                Else
                
                HighestIncr = HighestIncr
                HighestDecr = HighestDecr
                HighestVol = HighestVol
                
                End If
            'Enter summary data into cells
            ws.Cells(2, 17).Value = Format(HighestIncr, "Percent")
            ws.Cells(3, 17).Value = Format(HighestDecr, "Percent")
            ws.Cells(4, 17).Value = Format(HighestVol, "Scientific")
            
            Next CurrRow
            
            Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws


End Sub
