Attribute VB_Name = "Module1"
Sub alphabetical()

    Dim ticker As String
    Dim total_vol, open_value, close_value, yearly_change, percent_change As Double
    total_vol = 0
    open_value = 0
    close_value = 0

   Dim summary_row, ticker_count As Integer
    summary_row = 2
    ticker_count = 0

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
 
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "value"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
    
    
        ticker = Cells(i, 1).Value
        total_vol = total_vol + Cells(i, 7).Value
       
    
        Range("I" & summary_row).Value = ticker
        Range("L" & summary_row).Value = total_vol
        
        close_value = Cells(i, 6).Value
        yearly_change = close_value - open_value
        percent_change = (yearly_change / open_value)
        Range("j" & summary_row).Value = yearly_change

        
        If Range("j" & summary_row).Value > 0 Then
            Range("j" & summary_row).Interior.ColorIndex = 4
        Else
            Range("j" & summary_row).Interior.ColorIndex = 3
        End If
            
        Range("k" & summary_row).Value = percent_change
        Range("k" & summary_row).NumberFormat = "0.00"

        
        summary_row = summary_row + 1
        total_vol = 0
        ticker_count = 0
        
    Else
        total_vol = total_vol + Cells(i, 7).Value
        ticker_count = ticker_count + 1
        If ticker_count = 1 Then
            open_value = Cells(i, 3).Value
         End If
         
         Dim rng As Range
         Dim maxvalue As Variant
         Dim maxtotal As Variant
         Dim minvalue As Variant
         
       
         Set rng = Range("L" & summary_row)
         maxtotal = Application.WorksheetFunction.Max(rng)
         Range("P4").Value = maxtotal
         ticker = WorksheetFunction.Match(maxtotal, rng, 0)
         Range("O4").Value = Cells(ticker + 1, 9)
         
         
         Set rng = Range("K2:K91")
         maxvalue = Application.WorksheetFunction.Max(rng)
         Range("P2").Value = maxvalue
         ticker = WorksheetFunction.Match(maxvalue, rng, 0)
         Range("O2").Value = Cells(ticker + 1, 9)
         
         Set rng = Range("K2:K91")
         minvalue = Application.WorksheetFunction.Min(rng)
         Range("P3").Value = minvalue
         ticker = WorksheetFunction.Match(minvalue, rng, 0)
         Range("O3").Value = Cells(ticker + 1, 9)
        

        End If
    
    Next i



End Sub


