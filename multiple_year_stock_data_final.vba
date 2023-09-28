Sub stockanalysis()

'remember to activate all worksheets in the script before submitting the assignment

Dim ws As Worksheet
Dim select_index As Double
Dim first_row As Double
Dim select_row As Double
Dim last_row As Double
Dim year_open As Single
Dim year_close As Single
Dim volume As Double

    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        select_index = 2
        first_row = 2
        select_row = 2
        last_row = WorksheetFunction.CountA(ActiveSheet.Columns(1))
        volume = 0
        
        'Assign headers to the rows and columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume "
        
        'loop through all rows to identify unique tickers, place them in the new table
        For i = first_row To last_row
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i + 1, 1).Value
            If tickers <> tickers2 Then
                Cells(select_row, 9).Value = tickers
                select_row = select_row + 1
            End If
        Next i
        
        'next, loop through all rows and sum the volume if the ticker HAS NOT changed. when the ticker changes, reset the volume to 0. then move to the next row
        For i = first_row To last_row + 1
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i + 1, 1).Value
            If tickers = tickers2 And i > 2 Then
                volume = volume + Cells(i, 7).Value
            ElseIf i > 2 Then
                Cells(select_index, 12).Value = volume
                select_index = select_index + 1
                volume = 0
            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i
        
        'next, loop through all rows. if the the ticker symbol has changed, assign year_open. If the next ticker changes, assign year_close
        select_index = 2
        For i = first_row To last_row
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i + 1, 1).Value
            If tickers <> tickers2 Then
                year_close = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                year_open = Cells(i, 3).Value
            End If
            If year_open > 0 And year_close > 0 Then
                increase = year_close - year_open
                percent_increase = increase / year_open
                Cells(select_index, 10).Value = Format(increase, "#.00")
                Cells(select_index, 11).Value = FormatPercent(percent_increase)
                year_close = 0
                year_open = 0
                select_index = select_index + 1
            End If
        Next i
        
        'find min and max values
        max_per_yr = WorksheetFunction.Max(ActiveSheet.Columns("K"))
        min_per_yr = WorksheetFunction.Min(ActiveSheet.Columns("K"))
        max_vol_per_yr = WorksheetFunction.Max(ActiveSheet.Columns("L"))
        
        'format the first two values to percentages (leave volume as number) and assign them to the correct cells on each sheet
        Range("Q2").Value = FormatPercent(max_per_yr)
        Range("Q3").Value = FormatPercent(min_per_yr)
        Range("Q4").Value = Format(STANDARD, max_vol_per_yr)
        
        'next, loop through columns 11 and 12 to apply the corresponding ticker to the max/min/vol table
        For i = first_row To last_row
            If max_per_yr = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf min_per_yr = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
            ElseIf max_vol_per_yr = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
        Next i
        
        'Loop through column 10 and apply either green or red formatting
        For i = first_row To last_row
            If IsEmpty(Cells(i, 10).Value) Then Exit For
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 43
            Else
                Cells(i, 10).Interior.ColorIndex = 22
            End If
        Next i
    Next ws
        
                
End Sub
