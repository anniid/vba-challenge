Attribute VB_Name = "Module2"
Sub one_sheet_only()

'set all the dimensions
Dim total As Double
Dim i As Long
Dim j As Integer
Dim change As Single
Dim pChange As Single
Dim start As Long
Dim rowCount As Long

    'set titles
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    'set number format for the yearly change to match example picture
    Columns("J").NumberFormat = "0.00"
    'set variables for each sheet so that you always start from 0
    j = 0
    start = 2
    total = 0
    'get the rowcount for each sheet (they could be different) also, remember not to  this one
    rowCount = Cells(RoCount, 1).End(xlUp).Row
    'make sure that was right
    'MsgBox (rowCount)
    
    'if ticker changes, print results
    For i = 2 To rowCount
        If Cells(i + 1, 1) <> Cells(i, 1) Then
        'store the results in total
        total = total + Cells(i, 7).Value
        'but if the total is 0 that could be an issue
            If total = 0 Then
            'print results
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = 0
            Range("K" & 2 + j).Value = "%" & 0
            Range("L" & 2 + j).Value = 0
            Else
            'find first non-zero
            If Cells(start, 3) = 0 Then
                For Find = start To i
                    If Cells(Find, 3).Value <> 0 Then
                        start = Find
                        Exit For
                    End If
                Next Find
            End If
        
            'calculate changes
            change = (Cells(i, 6) - Cells(start, 3))
            pChange = Round((change / Cells(start, 3) * 100), 2)
            
            'start to the next stock ticker
            start = i + 1
            
            'print all the results so far: ticker, change, pchange, & total
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = change
            Range("K" & 2 + j).Value = "%" & pChange
            Range("L" & 2 + j).Value = total
            
            'change the colors for pos and neg changes
            If change > 0 Then
                Range("J" & 2 + j).Interior.ColorIndex = 4
            ElseIf change < 0 Then
                Range("J" & 2 + j).Interior.ColorIndex = 3
            Else
                Range("J" & 2 + j).Interior.ColorIndex = 0
            End If
            End If
            'reset the total and change so that we start again from 0 each time
            total = 0
            change = 0
            'move to next row in new table(j)
            j = j + 1
        Else
            total = total + Cells(i, 7).Value
        End If
        Next i
'moving on to the next table now for greatest %s and total
'find max % increase & print in new new table
Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
'find (min) max % decrease and print in new new table
Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
'find max total and print in new new table
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

'don't use the header row, find the position to put in the ticker in the new new table
inc = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
dec = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
vol = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

'print the tickers for the new new table
Range("P2") = Cells(inc + 1, 9)
Range("P3") = Cells(dec + 1, 9)
Range("P4") = Cells(vol + 1, 9)

'autofit the columns, because it's irritating.
Columns("I:Q").EntireColumn.AutoFit
                               

MsgBox ("Whew. That was just one sheet!")

End Sub
