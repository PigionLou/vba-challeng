Attribute VB_Name = "Mod_StockTracker1"
Sub StockTracker()

    Dim Ticker As String
    Dim opener, closer, volume, change, percentchange As Double
    Dim Last, i, m, lastsheet, sheetindex As Integer
    
    'determine last sheet
    lastsheet = Sheets.Count
    
    
    'iterate through each sheet
    For sheetindex = 1 To lastsheet
        
        With Sheets(sheetindex)
        'set initial values
            Last = .Cells(1, 1).End(xlDown).Row
            volume = 0
            m = 2
            
            'header formatting
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "% Change"
            .Cells(1, 12).Value = "Total Stock Volume"
            
            'for each row add the openeing, closing, and volume.
            For i = 2 To Last
                
                
                Ticker = .Cells(i, 1).Value
                volume = volume + .Cells(i, 7).Value
                
                'Store 1st opening value per ticker symbol
                If .Cells(i, 1).Value <> .Cells(i - 1, 1).Value Then
                    opener = .Cells(i, 3).Value
                
                'If the next ticker symbol is different calculate the difference and % difference for opening and closing and print the values to cells
                ElseIf .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
                    
                    closer = .Cells(i, 6).Value
                    change = closer - opener
                    
                    'error handler in case opening price is 0
                    If opener <> 0 Then
                        percentchange = (closer - opener) / opener
                    Else
                        percentchange = 1
                    End If
                    
                    'print values to cells
                    .Cells(m, 9).Value = Ticker
                    .Cells(m, 10).Value = change
                    .Cells(m, 11).Value = FormatPercent(percentchange)
                    .Cells(m, 12).Value = volume
        
                    'color formatting for yearly change
                    If change < 0 Then
                        .Cells(m, 10).Interior.Color = vbRed
                    Else
                        .Cells(m, 10).Interior.Color = vbGreen
                    End If
                    
        
                                
                    'add to summary counter and clear variables
                    m = m + 1
                    opener = 0
                    closer = 0
                    volume = 0
                    
                End If
                
            Next i
        
        .Range("A:P").Columns.AutoFit
        
        End With
    
    Next sheetindex
    

End Sub

