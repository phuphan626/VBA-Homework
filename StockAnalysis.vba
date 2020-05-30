Attribute VB_Name = "Module1"
'In order to complete this homework, I used solutions found on github from "kreitlerj" user.
'Link to his repository https://github.com/kreitlerj/VBA-Stock-Analysis/blob/master/StockAnalysis.vba
'I used his solution for guidance and referrences, so I will include my thought process just to clarify that
'I understand and did the assignment, not just copy and paste.
'Declare the subroutine
Sub StockAnalysis()
'Declare ws variable and loop through each worksheet
Dim ws As Worksheet
For Each ws In Worksheets
'Activate the current worksheet
    ws.Activate
'Add heading
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
'Declare all variables to hold values
    Dim ticker_name As String
    Dim volume As Double
    volume = 0
'Declare the column and row,so they will iterate, the code will be more dynamic
    Dim column As Integer
    column = 1
    Dim row As Long
    row = 2
'Declare delta as the yearly change variable
    Dim delta As Double
    delta = 0
    Dim percentage_change As Double
    percentage_change = 0
'Declare oprice and cprice for open price and close price, also declare the lastrow varaiable starting at the A1 column
    Dim oprice As Double
    oprice = 0
    Dim cprice As Double
    cprive = 0
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
'Conditional For loop to compare ticker symbol, then assign year open and total volume
    For i = 2 To lastrow
        'check for the next cell still the same with previous cell
        If ws.Cells(i, column).Value <> ws.Cells(i - 1, column).Value Then
                oprice = ws.Cells(i, 3).Value
            End If
                'Add volume up when found the different next cell
                volume = volume + ws.Cells(i, 7)
                
            If ws.Cells(i, column).Value <> ws.Cells(i + 1, column).Value Then
                ws.Cells(row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(row, 12).Value = volume
                cprice = ws.Cells(i, 6).Value
                delta = cprice - oprice
                ws.Cells(row, 10).Value = delta
                'using color index to identify the yearly change with 4 is green, 3 is red
                If delta > 0 Then
                ws.Cells(row, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(row, 10).Interior.ColorIndex = 3
                End If
                'This is to check the percentage change with formula (delta/oprice)-1=(close_price-open_price)/open_price
                If oprice = 0 And cprice = 0 Then
                percentage_change = 0
                ws.Cells(row, 11).Value = percentage_change
                ws.Cells(row, 11).NumberFormat = "0.00%"
                'if the open price is 0 and close price is not 0, so it will be infinite increase, I tried to but Null in, but it did not work,
                'so I replaced the percentage change with 1 after solving the formula as percentage change=infititive -1, and using NumberFormat function to get %
                ElseIf oprice = 0 And cprice <> 0 Then
                percentage_change = 1
                ws.Cells(row, 11).Value = percentage_change
                ws.Cells(row, 11).NumberFormat = "0.00%"
                Else
                percentage_change = delta / oprice
                ws.Cells(row, 11).Value = percentage_change
                ws.Cells(row, 11).NumberFormat = "0.00%"
                End If
                'After finding and adding up values, increase the row by 1 so it will increase one row in the summary table,
                'and reset all variable to 0
                row = row + 1
                volume = 0
                oprice = 0
                cprice = 0
                delta = 0
                percentage_change = 0
                
            End If
    Next i

'Create the performance table. this is to give credit to "kreitlerj"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).row
    
'declare best and worst stock and stock volume
    Dim best_stock As String
    Dim best_value As Double
    best_value = ws.Cells(2, 11).Value
    
    Dim worst_stock As String
    Dim worst_value As Double
    worst_value = ws.Cells(2, 11).Value
    
    Dim most_vol_stock As String
    Dim most_vol_value As Double
    most_vol_value = ws.Cells(2, 12).Value
    
    For o = 2 To lastrow
        If ws.Cells(o, 11).Value > best_value Then
        best_value = ws.Cells(o, 11).Value
        best_stock = ws.Cells(o, 9).Value
        End If
        If ws.Cells(o, 11).Value < worst_value Then
        worst_value = ws.Cells(o, 11).Value
        worst_stock = ws.Cells(o, 9).Value
        End If
        If ws.Cells(o, 12).Value > most_vol_value Then
        most_vol_value = ws.Cells(o, 12).Value
        most_vol_stock = ws.Cells(o, 9).Value
        End If
        'Move all data to performance table
        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit
    Next o
Next ws

End Sub

