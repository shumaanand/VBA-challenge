Attribute VB_Name = "Module1"
Sub Stocks():

'Loop through all worksheets
For Each ws In Worksheets

    ' Insert Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ' Define all Ticker Variables and set Default/Basline Variables
    Dim TickerName As String
    Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
    
    'Define all other Variable and set Dafault/Basline Variable
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim YearlyChange As Double
    
    Dim PreviousAmount As Long
        PreviousAmount = 2
    
    Dim PercentChange As Double
    
    Dim GreatestIncrease As Double
        GreatestIncrease = 0
    
    Dim GreatestDecrease As Double
        GreatestDecrease = 0
    
    Dim LastRowValue As Long
    
    Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0


    ' Set Summary Table for location/keep track of the location for each ticker name
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
        
    ' Determine the Last Row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all Tickers
    For i = 2 To LastRow

    ' Add to Ticker Total Volume
    TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
    
    ' Check if we are still within the same ticker name, and if it is not. . .
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Set TickerName data
        TickerName = ws.Cells(i, 1).Value
        ' Print the TickerName in the Summary Table
        ws.Range("I" & SummaryTableRow).Value = TickerName
        
        ' Print the Ticker Total Amount to the Summary Table
        ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
        ' Reset Ticker Total
        TotalTickerVolume = 0
        
    ' Set Yearly Open, Yearly Close and Yearly Change Name
    YearlyOpen = ws.Range("C" & PreviousAmount)
    YearlyClose = ws.Range("F" & i)
    YearlyChange = YearlyClose - YearlyOpen
    'Print location
    ws.Range("J" & SummaryTableRow).Value = YearlyChange

    ' Determine Percent Change
    If YearlyOpen = 0 Then
        PercentChange = 0
    Else
        YearlyOpen = ws.Range("C" & PreviousAmount)
        PercentChange = YearlyChange / YearlyOpen
    End If
    
    ' Format Double To Include % Symbol And Two Decimal Places
    ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
    ws.Range("K" & SummaryTableRow).Value = PercentChange

    ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
    If ws.Range("J" & SummaryTableRow).Value >= 0 Then
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
    Else
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
    End If
            
    ' Add One To The Summary Table Row
    SummaryTableRow = SummaryTableRow + 1
    PreviousAmount = i + 1
        
            End If
        
     Next i
     ' Format Table Columns To Auto Fit
        ws.Columns("I:Q").AutoFit
     
   Next ws
    
End Sub





