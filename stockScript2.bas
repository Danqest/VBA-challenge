Attribute VB_Name = "Module13"
Sub StockScript()
    
    ' Declare loop to cycle through each worksheet
    Dim ws As Worksheet
    
    ' Loop through worksheets
    For Each ws In Worksheets
        ws.Activate
    
        ' Prepare titles for summary columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
    
        ' Get last row of data in Column A for dynamic scalability in script
        Dim LastRow As LongLong
        LastRow = Range("A2").End(xlDown).Row
        
        ' Select date column data and store first and last date values as variables for reference later
        Dim StartDate As Long
        Dim EndDate As Long
        StartDate = Range("B2").Value
        EndDate = Range("B" & LastRow).Value
        
        ' Declare storage variables
        Dim SelectedTicker As String
        Dim VolumeCount As LongLong
        Dim FirstPrice As Double
        Dim LastPrice As Double
        
        ' Format Yearly Change column to two decimal places
        Range("J:J").NumberFormat = "0.00"
        
        
        
        ' Select ticker column data and return unique ticker values in summary ticker column
        Set TickerRng = Range("A2:A" & LastRow)
        Set TickerUnique = Range("I2")
        
        Dim tickerCount As Double
        tickerCount = 2
        
        ' Initialize VolumeCount
        VolumeCount = 0
        
        ' Prefill first ticker for loop below
        Range("I2").Value = Range("A2").Value
        
        ' Loop through all of main dataset to lastrow + 1 so the final ticker's data summary fills
        For a = 2 To LastRow + 1
            ' If the main ticker DOES NOT match the summary ticker
            If Range("A" & a) <> Range("I" & tickerCount) Then
                ' Fill and calculate stored summary data
                Cells(tickerCount, 12).Value = VolumeCount
                Cells(tickerCount, 10).Value = (EndPrice - FirstPrice)
                Cells(tickerCount, 11).Value = FormatPercent((EndPrice - FirstPrice) / FirstPrice)
        
                ' Format Yearly Change cell color based on if positive or negative number
                If Cells(tickerCount, 10).Value < 0 Then
                    Cells(tickerCount, 10).Interior.Color = vbRed
                End If
                If Cells(tickerCount, 10).Value > 0 Then
                    Cells(tickerCount, 10).Interior.Color = vbGreen
                End If
                
                ' Fill in next summary ticker, increment tickerCount, reset VolumeCount
                Cells(tickerCount + 1, 9).Value = Cells(a, 1).Value
                tickerCount = tickerCount + 1
                VolumeCount = 0
            End If
            
            ' If the main ticker DOES match the summary ticker
            ' store aggregated VolumeCount, open price on first day, and close on last day
            If Range("A" & a) = Range("I" & tickerCount) Then
                    VolumeCount = VolumeCount + Cells(a, 7)
                    
                    If Cells(a, 2).Value = StartDate Then
                        FirstPrice = Cells(a, 3).Value
                    End If
                    
                    If Cells(a, 2).Value = EndDate Then
                        EndPrice = Cells(a, 6).Value
                    End If
                    
            End If
               
        Next a
        
        
        ' Select unique ticker column summary data and store last row data as variable
        Dim SummaryTickerLastRow As Integer
        SummaryTickerLastRow = Range("I2").End(xlDown).Row
        
        
        ' Prepare Greatest %Increase/%Decrease/Total Volume columns next to summary data
        Range("P1").Value = "Ticker"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("Q1").Value = "Value"
        
        
        ' Fill Value columns prepared above
        Range("Q2").Value = FormatPercent(WorksheetFunction.Max(Range("K2:K" & SummaryTickerLastRow)))
        Range("Q3").Value = FormatPercent(WorksheetFunction.Min(Range("K2:K" & SummaryTickerLastRow)))
        Range("Q4").Value = WorksheetFunction.Max(Range("L2:K" & SummaryTickerLastRow))
        
        ' Loop through summary data and retrieve Tickers
        For k = 2 To SummaryTickerLastRow
            If Cells(k, 11).Value = Range("Q2").Value Then
                Range("P2").Value = Range("I" & k).Value
            End If
            
            If Cells(k, 11).Value = Range("Q3").Value Then
                Range("P3").Value = Range("I" & k).Value
            End If
            
            If Cells(k, 12).Value = Range("Q4").Value Then
                Range("P4").Value = Range("I" & k).Value
            End If
            
        Next k
    
    ' Move to next worksheet
    Next
    


End Sub
