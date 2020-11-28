Sub TickerProcessorAllWorksheets()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        
        ' Created a Variable to tickerName, rowCount
        Dim WorksheetName As String

        'declare worksheet rowCount
        Dim rowCount As Long: rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
           
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        'MsgBox (WorksheetName)
        
        'declare and define summaryTableRow row counter
        Dim summaryTableRow As Integer: summaryTableRow = 2
               
        'declare j row interator
        Dim j As Long
        
        Dim isOpeningTicker As Boolean: isOpeningTicker = True
                        
        Dim tickerOpenValue As Double
        Dim tickerCloseValue As Double
        Dim tickerChange As Double
        Dim tickerPercentChange As Double
        Dim tickerVolumeTotal As Double
        Dim tickerName As String
        Dim maxPercentChange As Double: maxPercentChange = 0
        Dim minPercentChange As Double: minPercentChange = 0
        Dim maxVolumeTotal As Double: maxVolumeTotal = 0
        
        'set summary headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
             
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Stock Volume"
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
               
        'start nested for loop to iterate
        For j = 2 To rowCount
            ticker = ws.Cells(j, 1).Value
            
            ' ticker open value if first instance of ticker
            If isOpeningTicker = True Then
                tickerOpenValue = ws.Cells(j, 3).Value
            End If
                                
            'check if they're the same if they're not the same
            If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                'set ticker name
                tickerName = ws.Cells(j, 1).Value
                                               
                'add to tickerVolumeTotal
                tickerVolumeTotal = tickerVolumeTotal + ws.Cells(j, 7).Value
                
                'find yearly change
                tickerCloseValue = ws.Cells(j, 6).Value
                tickerChange = tickerCloseValue - tickerOpenValue
                
                'handle if tickerOpenValue = 0 (cannot divide by zero, set ticker percent change to 1
                If tickerOpenValue <> 0 Then
                    '% increase = Increase ÷ Original Number
                    tickerPercentChange = (tickerChange / tickerOpenValue)
                ElseIf tickerChange <> 0 Then
                    tickerPercentChange = 1
                Else 'tickerChange = 0
                    tickerPercentChange = 0
                End If
                                                      
                'print tickerChange in summary table
                ws.Range("I" & summaryTableRow).Value = tickerName
                
                'print tickerPercentChange in summary table
                ws.Range("K" & summaryTableRow).Value = FormatPercent(tickerPercentChange, 2)
                
                'set maxPercentChange
                maxPercentChange = ws.Range("P2").Value
                
                'set/print maxPercentChange in summary table
                If tickerPercentChange > maxPercentChange And tickerPercentChange > 0 Then
                    ws.Range("P2").Value = FormatPercent(tickerPercentChange, 2)
                    ws.Range("O2").Value = tickerName
                End If
                
                'set minPercentChange
                minPercentChange = ws.Range("P3").Value
                
                'set/print minPercentChange in summary table
                If tickerPercentChange < minPercentChange And tickerPercentChange < 0 Then
                    ws.Range("P3").Value = FormatPercent(tickerPercentChange, 2)
                    ws.Range("O3").Value = tickerName
                End If
                
                
                'set maxVolumeTotal
                maxVolumeTotal = ws.Range("P4").Value
                
                'set/print maxVolumeTotal in summary table
                If tickerVolumeTotal > maxVolumeTotal Then
                    ws.Range("P4").Value = tickerVolumeTotal
                    ws.Range("O4").Value = tickerName
                End If
                
                
                                                                   
                If tickerChange < 0 Then
                    'set tickerchange and ticker percent change color: dark red font / light red interior fill
                    ws.Range("J" & summaryTableRow).Font.Color = RGB(200, 50, 50)
                    ws.Range("J" & summaryTableRow).Interior.Color = RGB(242, 194, 194)
                    
                    ws.Range("K" & summaryTableRow).Font.Color = RGB(200, 50, 50)
                    ws.Range("K" & summaryTableRow).Interior.Color = RGB(242, 194, 194)
                ElseIf tickerChange > 0 Then
                    'set tickerchange and ticker percent change color: dark green font / light green interior fill
                    ws.Range("J" & summaryTableRow).Font.Color = RGB(4, 105, 28)
                    ws.Range("J" & summaryTableRow).Interior.Color = RGB(194, 242, 204)
                    
                    ws.Range("K" & summaryTableRow).Font.Color = RGB(4, 105, 28)
                    ws.Range("K" & summaryTableRow).Interior.Color = RGB(194, 242, 204)
                End If
                                                
                'print tickerChange $ in summary table
                ws.Range("J" & summaryTableRow).Value = FormatCurrency(tickerChange, 2)
                
                'print tickerVolumeTotal in summary table
                ws.Range("L" & summaryTableRow).Value = tickerVolumeTotal
                
                'increment summaryTableRow counter
                summaryTableRow = summaryTableRow + 1
                
                'reset tickerVolumeTotal
                tickerVolumeTotal = 0
                
                'reset tickerStartValue
                tickerStartValue = 0
                
                'reset opening Ticker
                isOpeningTicker = True
                                
            Else
                'add vol total (same ticker number)
                tickerVolumeTotal = tickerVolumeTotal + ws.Cells(j, 7).Value
                
                'set opening Ticker False
                isOpeningTicker = False
                                            
            End If
        Next j
        
        'reset rowCount for next worksheet
        rowCount = 0
        
        'reset rowCount for next worksheet
        summaryTableRow = 2
        
        'reset max/min summary variables
        maxPercentChange = 0
        minPercentChange = 0
    Next ws
    
MsgBox ("Ticket Processor Completed for all worksheets")

End Sub


Sub TickerProcessor()

'Objectives:
'DONE - List unique ticker symbol.
'DONE - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'DONE - The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'DONE - The total stock volume of the stock.
'DONE - conditional formatting that will highlight positive change in green and negative change in red.
  
    'declare and define summaryTableRow row counter
    Dim summaryTableRow As Double: summaryTableRow = 2
    
    'declare worksheet rowCount
    Dim rowCount As Long: rowCount = Cells(Rows.Count, 1).End(xlUp).Row
    
    'declare j row interator
    Dim j As Long
    
    Dim isOpeningTicker As Boolean: isOpeningTicker = True
            
    'Dim worksheetName As String
    Dim tickerOpenValue As Double
    Dim tickerCloseValue As Double
    Dim tickerChange As Double
    Dim tickerPercentChange As Double
    Dim tickerVolumeTotal As Double
    Dim tickerName As String
    Dim maxPercentChange As Double: maxPercentChange = 0
    Dim minPercentChange As Double: minPercentChange = 0
    Dim maxVolumeTotal As Double: maxVolumeTotal = 0
    
    'set summary headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
         
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Stock Volume"
    
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
            
    'start nested for loop to iterate
    For j = 2 To rowCount
        ticker = Cells(j, 1).Value
        
        ' ticker open value if first instance of ticker
        If isOpeningTicker = True Then
            tickerOpenValue = Cells(j, 3).Value
        End If
                            
        'check if they're the same if they're not the same
        If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
            'set ticker name
            tickerName = Cells(j, 1).Value
                                           
            'add to tickerVolumeTotal
            tickerVolumeTotal = tickerVolumeTotal + Cells(j, 7).Value
            
            'find yearly change
            tickerCloseValue = Cells(j, 6).Value
            tickerChange = tickerCloseValue - tickerOpenValue
            
            'handle if tickerOpenValue = 0 (cannot divide by zero, set ticker percent change to 1
            If tickerOpenValue <> 0 Then
                '% increase = Increase ÷ Original Number
                tickerPercentChange = (tickerChange / tickerOpenValue)
            ElseIf tickerChange <> 0 Then
                tickerPercentChange = 1
            Else 'tickerChange = 0
                tickerPercentChange = 0
            End If
                                                  
            'print tickerChange in summary table
            Range("I" & summaryTableRow).Value = tickerName
            
            'print tickerPercentChange in summary table
            Range("K" & summaryTableRow).Value = FormatPercent(tickerPercentChange, 2)
            
            'set maxPercentChange
            maxPercentChange = Range("P2").Value
            
            'set/print maxPercentChange in summary table
            If tickerPercentChange > maxPercentChange And tickerPercentChange > 0 Then
                Range("P2").Value = FormatPercent(tickerPercentChange, 2)
                Range("O2").Value = tickerName
            End If
            
            'set minPercentChange
            minPercentChange = Range("P3").Value
            
            'set/print minPercentChange in summary table
            If tickerPercentChange < minPercentChange And tickerPercentChange < 0 Then
                Range("P3").Value = FormatPercent(tickerPercentChange, 2)
                Range("O3").Value = tickerName
            End If
            
            
            'set maxVolumeTotal
            maxVolumeTotal = Range("P4").Value
            
            'set/print maxVolumeTotal in summary table
            If tickerVolumeTotal > maxVolumeTotal Then
                Range("P4").Value = tickerVolumeTotal
                Range("O4").Value = tickerName
            End If
            
            
                                                               
            If tickerChange < 0 Then
                'set tickerchange and ticker percent change color: dark red font / light red interior fill
                Range("J" & summaryTableRow).Font.Color = RGB(200, 50, 50)
                Range("J" & summaryTableRow).Interior.Color = RGB(242, 194, 194)
                
                Range("K" & summaryTableRow).Font.Color = RGB(200, 50, 50)
                Range("K" & summaryTableRow).Interior.Color = RGB(242, 194, 194)
            ElseIf tickerChange > 0 Then
                'set tickerchange and ticker percent change color: dark green font / light green interior fill
                Range("J" & summaryTableRow).Font.Color = RGB(4, 105, 28)
                Range("J" & summaryTableRow).Interior.Color = RGB(194, 242, 204)
                
                Range("K" & summaryTableRow).Font.Color = RGB(4, 105, 28)
                Range("K" & summaryTableRow).Interior.Color = RGB(194, 242, 204)
            End If
                                            
            'print tickerChange $ in summary table
            Range("J" & summaryTableRow).Value = FormatCurrency(tickerChange, 2)
            
            'print tickerVolumeTotal in summary table
            Range("L" & summaryTableRow).Value = tickerVolumeTotal
            
            'increment summaryTableRow counter
            summaryTableRow = summaryTableRow + 1
            
            'reset tickerVolumeTotal
            tickerVolumeTotal = 0
            
            'reset tickerStartValue
            tickerStartValue = 0
            
            'reset opening Ticker
            isOpeningTicker = True
                            
        Else
            'add vol total (same ticker number)
            tickerVolumeTotal = tickerVolumeTotal + Cells(j, 7).Value
            
            'set opening Ticker False
            isOpeningTicker = False
                                        
        End If
    Next j
            
    'reset rowCount
    rowCount = 0
    
    'reset rowCount
    summaryTableRow = 2
    
    'reset max/min summary variables
    maxPercentChange = 0
    minPercentChange = 0
        
MsgBox ("Ticket Processor Completed for this worksheet")

End Sub



Sub ClearTickerSummary()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        
        ' Created a Variable to tickerName, rowCount
        Dim WorksheetName As String

        'find row count
        Dim rowCount As Long: rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'clear contents from second row to last row rowCount
        ws.Range("I2:L" & rowCount).Value = ""
        
        'clear formatting from second row to last row rowCount
        ws.Range("I2:L" & rowCount).Font.Color = RGB(0, 0, 0)
        ws.Range("I2:L" & rowCount).Interior.Color = xlNone
        
        
        'clear summary table greatest increase/decrease greatest total volume
        ws.Range("O2:P4").Value = ""

    Next ws
End Sub


