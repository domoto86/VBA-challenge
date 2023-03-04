Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data():

For Each xz In Worksheets
    
    Dim i, LastRow As Long
    Dim Ticker As String
    Dim TableRow As Integer
    Dim TickerTotal, YearOpen, YearClose, YearChange, PercentChange, GreatIncrease, GreatDecrease, GreatTotal As Double 'Total Stock Volume, Year Open and Year Close
    
    xz.Range("I1").Value = "Ticker"
    xz.Range("J1").Value = "Yearly Change"
    xz.Range("K1").Value = "Percent Change"
    xz.Range("L1").Value = "Total Stock Volume"
    xz.Range("O2").Value = "Greatest % Increase"
    xz.Range("O3").Value = "Greatest % Decrease"
    xz.Range("O4").Value = "Greatest Total Volume"
    xz.Range("P1").Value = "Ticker"
    xz.Range("Q1").Value = "Value"
    
    TableRow = 2
    TickerTotal = 0
    YearOpen = 0
    YearClose = 0
    YearChange = 0
    PercentChange = 0
    
    LastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To LastRow
    
        If (xz.Cells(i + 1, "A").Value) <> (xz.Cells(i, "A").Value) Then
            
            Ticker = xz.Cells(i, "A").Value
            
            TickerTotal = TickerTotal + xz.Cells(i, "G").Value
            
            YearClose = xz.Cells(i, "F").Value
            
            YearChange = YearClose - YearOpen
            PercentChange = YearChange / YearOpen
            
            xz.Cells(TableRow, "I") = Ticker
            xz.Cells(TableRow, "L") = TickerTotal
            xz.Cells(TableRow, "J") = YearChange
            xz.Cells(TableRow, "K") = PercentChange
            
            TableRow = TableRow + 1
            
            TickerTotal = 0
            YearOpen = 0
            YearClose = 0
        
        ElseIf (xz.Cells(i - 1, "A").Value) <> (xz.Cells(i, "A").Value) Then
            
            YearOpen = xz.Cells(i, "C").Value
            
            TickerTotal = TickerTotal + xz.Cells(i, "G").Value
        
        Else
            
            TickerTotal = TickerTotal + xz.Cells(i, "G").Value
        
        End If
    
    Next i
    
    'Conditional Formatting
    
    LastRowSummary = ActiveSheet.Cells(Rows.Count, "I").End(xlUp).Row
    
    For j = 2 To LastRowSummary
    
        If xz.Cells(j, "J").Value > 0 Then
            
            xz.Cells(j, "J").Interior.ColorIndex = 4
        
        Else
            
            xz.Cells(j, "J").Interior.ColorIndex = 3
        
        End If
    
    Next j
    
    'Bonus part of the assignment
    
    GreatIncrease = 0
    GreatDecrease = 0
    GreatTotal = 0
    
    For k = 2 To LastRowSummary
    
        If xz.Cells(k, "K").Value > GreatIncrease Then
            
            GreatIncrease = xz.Cells(k, "K").Value
            xz.Range("Q2").Value = GreatIncrease
            xz.Range("P2").Value = xz.Cells(k, "I").Value
        
        End If
        
        If xz.Cells(k, "K").Value < GreatDecrease Then
            
            GreatDecrease = xz.Cells(k, "K").Value
            xz.Range("Q3").Value = GreatDecrease
            xz.Range("P3").Value = xz.Cells(k, "I").Value
        
        End If
        
        If xz.Cells(k, "L").Value > GreatTotal Then
            
            GreatTotal = xz.Cells(k, "L").Value
            xz.Range("Q4").Value = GreatTotal
            xz.Range("P4").Value = xz.Cells(k, "I").Value
        
        End If
    
    Next k
    
    'Format for Percentage and Autofit
    
    xz.Columns("K:K").NumberFormat = "0.00%"
    xz.Range("Q2:Q3").NumberFormat = "0.00%"
        
    xz.Columns("A:Q").AutoFit

Next xz
    
End Sub


