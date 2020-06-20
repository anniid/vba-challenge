Attribute VB_Name = "Module1"
Sub stock_summary()
'set all the dimensions
Dim total As Double
Dim i As Long
Dim j As Integer
Dim change As Single
Dim pChange As Single
Dim start As Long
Dim rowCount As Long
Dim ws As Worksheet

For Each ws In Worksheets

    'set titles
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    'set number format for the yearly change to match example picture
    ws.Columns("J").NumberFormat = "0.00"
    'set variables for each sheet so that you always start from 0
    j = 0
    start = 2
    total = 0
    'get the rowcount for each sheet (they could be different) also, remember not to ws. this one
    rowCount = Cells(Rows.Count, 1).End(xlUp).Row
    'make sure that was right
    'MsgBox (rowCount)
    
    'if ticker changes, print results
    For i = 2 To rowCount
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        'store the results in total
        total = total + ws.Cells(i, 7).Value
        'but if the total is 0 that could be an issue
            If total = 0 Then
            'print results
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = 0
            ws.Range("K" & 2 + j).Value = "%" & 0
            ws.Range("L" & 2 + j).Value = 0
            Else
            'find first non-zero
            If ws.Cells(start, 3) = 0 Then
                For Find = start To i
                    If ws.Cells(Find, 3).Value <> 0 Then
                        start = Find
                        Exit For
                    End If
                Next Find
            End If
        
            'calculate changes
            change = (ws.Cells(i, 6) - ws.Cells(start, 3))
            pChange = Round((change / ws.Cells(start, 3) * 100), 2)
            
            'start to the next stock ticker
            start = i + 1
            
            'print all the results so far: ticker, change, pchange, & total
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = change
            ws.Range("K" & 2 + j).Value = "%" & pChange
            ws.Range("L" & 2 + j).Value = total
            
            'change the colors for pos and neg changes
            If change > 0 Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            ElseIf change < 0 Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            Else
                ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            End If
            End If
            'reset the total and change so that we start again from 0 each time
            total = 0
            change = 0
            'move to next row in new table(j)
            j = j + 1
        Else
            total = total + ws.Cells(i, 7).Value
        End If
        Next i
'moving on to the next table now for greatest %s and total
'find max % increase & print in new new table
ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
'find (min) max % decrease and print in new new table
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
'find max total and print in new new table
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

'don't use the header row, find the position to put in the ticker in the new new table
inc = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
dec = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
vol = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

'print the tickers for the new new table
ws.Range("P2") = ws.Cells(inc + 1, 9)
ws.Range("P3") = ws.Cells(dec + 1, 9)
ws.Range("P4") = ws.Cells(vol + 1, 9)

'autofit the columns, because it's irritating.
ws.Columns("I:Q").EntireColumn.AutoFit
                               
'next worksheet
Next ws

MsgBox ("Whew. All done.")

End Sub
