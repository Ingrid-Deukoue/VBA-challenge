Attribute VB_Name = "Module1"
Sub VBA_Stock_Homework()

'Define all headers & variables

For Each ws In Worksheets
    'colums header
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = ""
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Variables
    
    Dim TickerName As String
    Dim LastRow As Long
    Dim Ticker As String
    Dim TotalTickerVolume As Double
    TotalTickerVolume = 0
    Dim PreviousValue As Long
    PreviousValue = 2
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim LastRowValue As Long
    Dim TotalStockVolume As Long
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = 0
    
    
        
    'last row calculation
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    For i = 2 To LastRow
    
    TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Total Stock Volume calculation
        
        TickerName = ws.Cells(i, 1).Value
        ws.Range("I" & SummaryTableRow).Value = TickerName
        ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
        TotalTickerVolume = 0
        
        
        'Yearly Open, Close and change
        
        YearOpen = ws.Range("C" & PreviousValue)
        YearClose = ws.Range("F" & i)
        YearlyChange = (YearClose - YearOpen)
        ws.Range("J" & SummaryTableRow).Value = PercentChange
    
    If YearOpen = 0 Then
        PercentChange = 0
    Else
        YearOpen = ws.Range("C" & PreviousValue)
        PercentChange = YearlyChange / YearOpen
    End If
    
    'Percentage Formatting
    ws.Range("K" & SummaryTableRow).Value = "0.00%"
    ws.Range("K" & SummaryTableRow).Value = PercentChange
    
    'Highlight formatting
    If ws.Range("J" & SummaryTableRow).Value >= 0 Then
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
    Else
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
    
    End If
    
    SummaryTableRow = SummaryTableRow + 1
    PreviousValue = i + 1
       
    End If
    Next i
    
    'greatest % Increase/Decrease & Volume
    'Determine Last Row
    
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To LastRow
        If ws.Range("K" & i) > ws.Range("Q2").Value Then
        ws.Range("Q2").Value = ws.Range("K" & i)
        ws.Range("P2").Value = ws.Range("I" & i)
    End If
    
        If ws.Range("K" & i) < ws.Range("Q3").Value Then
        ws.Range("Q3").Value = ws.Range("K" & i)
        ws.Range("P3").Value = ws.Range("I" & i)
    End If
    
    If ws.Range("L" & i) > ws.Range("Q4").Value Then
        ws.Range("Q4").Value = ws.Range("L" & i)
        ws.Range("P4").Value = ws.Range("I" & i)
    End If
    
Next i

    'Formatting
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
        
        
Next ws

End Sub
