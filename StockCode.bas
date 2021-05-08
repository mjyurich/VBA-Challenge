Attribute VB_Name = "StockCode"
Sub VBAChallenge()

For Each ws In Worksheets

'Set variables needed and total stock
Dim Ticker_Name As String
Dim total_stock As Double
total_stock = 0

'For location of each ticker symbol
Dim SummaryTableRow As Long
SummaryTableRow = 2

'Set Variables for yearly change
Dim YearlyOpen As Double
Dim YearlyClose As Double
Dim YearlyChange As Double
Dim PreviousAmount As Long
PreviousAmount = 2

'Set Variables for Percent Change
Dim PercentChange As Double

'Create titles for categories/Column Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Determine the Last Row
Dim Lastrow As Long
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Start for loop on stock volume
For i = 2 To Lastrow

    'Add to stock total
    total_stock = total_stock + ws.Cells(i, 7).Value

    'Make sure we are in the same ticker name
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set Ticker
        Ticker_Name = Cells(i, 1).Value
        'Print Ticker Symbol in the column
        ws.Cells(SummaryTableRow, 9).Value = Ticker_Name
        'Show total stock in the info table
        ws.Cells(SummaryTableRow, 12).Value = total_stock
        'stock total at 0
        total_stock = 0
        
        'Set Yearly Change Name, yearly open and yearly close
        YearlyOpen = ws.Range("C" & PreviousAmount).Value
        YearlyClose = ws.Range("F" & i).Value
        YearlyChange = YearlyClose - YearlyOpen
        ws.Range("J" & SummaryTableRow).Value = YearlyChange
        
        'Determine Percent Change
        If YearlyOpen = 0 Then
            
            PercentChange = 0
        
        Else
            
            'Percent Change formula and place in Summary Table
            YearlOpen = ws.Range("C" & PreviousAmount).Value
            PercentChange = (YearlyChange / YearlyOpen)
            
        End If
        
        'Put Percent Change category in table, Convert to percent
        ws.Cells(SummaryTableRow, 11).Value = PercentChange
        ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
        
        'Condition formatting
        If ws.Range("J" & SummaryTableRow).Value > 0 Then
            
            'If higher that 0, green
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        
        ElseIf ws.Range("J" & SummaryTableRow).Value < 0 Then
            
            'If lower than 0, red
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            
        Else
        
            'If 0, then leave white
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 0
        
        End If
        
        'Add one to each Summary Table Row
        SummaryTableRow = SummaryTableRow + 1
        PreviousAmount = i + 1
        
    End If
    
Next i

'Variables for bonus question
Dim greatestincrease As Double
greatestincrease = 0
Dim greatestdecrease As Double
greatestdecrease = 0
Dim greatestvolume As Double
greatestvolume = 0

'For loop to start bonus question
For j = 2 To Lastrow

    'Find the value in Summary table with greatest increase
    'Put ticker number, convert to percent
    If ws.Range("k" & j).Value > greatestincrease Then
    
        greatestincrease = ws.Range("k" & j).Value
        ws.Range("Q2").Value = greatestincrease
        ws.Range("P2").Value = ws.Cells(j, 9).Value
        ws.Range("Q2").NumberFormat = "0.00%"
        
    End If
        
    'Find the value in Summary table with greatest decrease
    'Put ticker number, convert to percent
    If ws.Range("k" & j).Value < greatestdecrease Then
    
        greatestdecrease = ws.Range("k" & j).Value
        ws.Range("Q3").Value = greatestdecrease
        ws.Range("P3").Value = ws.Cells(j, 9).Value
        ws.Range("Q3").NumberFormat = "0.00%"
        
    End If
    
    'Find the greatest stock volume in Summary table
    'Find corresponding ticker symbol
    If ws.Range("L" & j).Value > greatestvolume Then
    
        greatestvolume = ws.Range("L" & j).Value
        ws.Range("Q4").Value = greatestvolume
        ws.Range("P4").Value = ws.Cells(j, 9).Value
    
    End If
    
Next j

Next ws
    
End Sub
