Attribute VB_Name = "Module1"
Sub stockprices()

    'Looping/indexing variables
    Dim sheet As Worksheet ' sheet variable
    Dim i As Long ' variable for iteration
    Dim j As Long ' variable for iteration
    Dim OutputRow As Long 'Variable to track rows of output columns
    OutputRow = 2
    Dim counter As Integer 'for open do loop
    
    
    '---------------------------------------------------
    'Creating actual variables
    Dim Ticker_Name As String 'Ticker name variable
    Dim YearOpenPrice As Double 'Yearly open price variable
    Dim YearClosingPrice As Double 'Yearly closing price variable
    Dim YearlyChange As Double 'Yearly price change variable
    Dim PercentChange As Double 'Yearly percentage change variable
    Dim TotalStockVolume As Double 'Variable for total stock volume for each ticker
    
    ' variables for second output table
    Dim maxincrease_ticker As String
    Dim maxincrease As Double
    Dim maxdecrease_ticker As String
    Dim maxdecrease As Double
    Dim maxvolume_ticker As String
    Dim maxvolume As Double
    
    TotalStockVolume = 0
    
    '---------------------------------------------------
    'Set up the output columns & worksheet stuff
    For Each sheet In Worksheets
    
    ' Resetting output row
        OutputRow = 2
        
    ' Variable to count rows of each worksheet for iteration
        Dim RowsCount As Long
        RowsCount = sheet.Range("A1").End(xlDown).Row
        
        Dim LastOutputRow As Integer
        
    ' Insert Ticker, Yearly Change, Percentage Change, and Total Stock Volume cols
        sheet.Range("I1").EntireColumn.Insert
        sheet.Cells(1, 9).Value = "Ticker"
        
        sheet.Range("J1").EntireColumn.Insert
        sheet.Cells(1, 10).Value = "Yearly Change"
        
        sheet.Range("K1").EntireColumn.Insert
        sheet.Cells(1, 11).Value = "Percent Change"
        
        sheet.Range("L1").EntireColumn.Insert
        sheet.Cells(1, 12).Value = "Total Stock Volume"
        
    ' Insert second output table
        sheet.Range("P1").EntireColumn.Insert
        sheet.Cells(1, 16).Value = "Ticker"
        
        sheet.Range("Q1").EntireColumn.Insert
        sheet.Cells(1, 17).Value = "Value"
        
        sheet.Cells(2, 15).Value = "Greatest % Increase"
        sheet.Cells(3, 15).Value = "Greatest % Decrease"
        sheet.Cells(4, 15).Value = "Greatest Total Volume"
    
    '---------------------------------------------------------------
    ' Looping time
        For i = 2 To RowsCount
        
            'If cells/tickers in following row are different
            If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
                
                ' Filling in ticker names
                Ticker_Name = sheet.Cells(i, 1).Value
                sheet.Range("I" & OutputRow).Value = Ticker_Name
                
                'Calculating closing price
                YearClosingPrice = sheet.Cells(i, 6).Value
                
                'Calculate YearlyChange
                YearlyChange = YearClosingPrice - YearOpenPrice
                sheet.Range("J" & OutputRow).Value = YearlyChange
                
                'Calculate PercentChange
                PercentChange = (YearlyChange / YearOpenPrice)
                sheet.Range("K" & OutputRow).NumberFormat = "0.00%"
                sheet.Range("K" & OutputRow).Value = PercentChange
                
                ' Calculate and fill in total stock volume
                TotalStockVolume = TotalStockVolume + sheet.Cells(i, 7).Value
                sheet.Range("L" & OutputRow).Value = TotalStockVolume
                
                'finding max volume, percent increase and decrease
                If PercentChange > maxincrease Then
                    maxincrease = PercentChange
                    maxincrease_ticker = Ticker_Name
                ElseIf PercentChange < maxdecrease Then
                    maxdecrease = PercentChange
                    maxdecrease_ticker = Ticker_Name
                End If
                
                If TotalStockVolume > maxvolume Then
                    maxvolume = TotalStockVolume
                    
                    maxvolume_ticker = Ticker_Name
                End If
                
                ' Add one to output table row
                OutputRow = OutputRow + 1
                
                'reset stock volume
                TotalStockVolume = 0
                
                'reset counter for yearly open
                counter = 0
                
            ' If the ticker cells are the same
            Else
                'Add to total volume
                 TotalStockVolume = TotalStockVolume + sheet.Cells(i, 7).Value
                 
                 ' Calculating opening price for each ticker
                 Do Until counter = 1
                    YearOpenPrice = sheet.Cells(i, 3).Value
                    YearOpenPrice = YearOpenPrice
                    counter = 1
                 Loop
                
            End If
            
        Next i
        
        ' Outputing max volume, percent increase and decrease
        sheet.Range("P2").Value = maxincrease_ticker
        sheet.Range("Q2").Value = maxincrease
        
        sheet.Range("P3").Value = maxdecrease_ticker
        sheet.Range("Q3").Value = maxdecrease
        
        sheet.Range("P4").Value = maxvolume_ticker
        sheet.Range("Q4").Value = maxvolume
        sheet.Range("Q4").NumberFormat = "General"
        
        'Color coding columns
        For i = 2 To OutputRow
            If sheet.Cells(i, 10).Value >= 0 Then
                sheet.Cells(i, 10).Interior.ColorIndex = 4 'Green
            Else
                sheet.Cells(i, 10).Interior.ColorIndex = 3 'Red
            End If
            
            If sheet.Cells(i, 11).Value >= 0 Then
                sheet.Cells(i, 11).Interior.ColorIndex = 4 'Green
            Else
                sheet.Cells(i, 11).Interior.ColorIndex = 3 'Red
            End If
            
        Next i
        
    Next sheet

End Sub

