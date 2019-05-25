Sub stockdata()
'Create a script that will loop through one year of stock data for each run and
'return the total volume each stock had over that year.

Dim ticker_name As String

Dim ticker_total As Double
ticker_total = 0

Dim year_change As Double

Dim percent_change As Double

Dim open_value As Double
Dim close_value As Double
        
'Loop through each worksheet
For Each ws In Worksheets

'Keep track of the location for each ticker in the summary table.
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
      
' Add the word Ticker to the Header
ws.Cells(1, 9).Value = "Ticker"

' Add the words Yearly Change to the Header
ws.Cells(1, 10).Value = "Yearly Change"

'Add the words Percent Change to the Header
ws.Cells(1, 11).Value = "Percent Change"
        
' Add the word Total Stock Volume to the Header
ws.Cells(1, 12).Value = "Total Stock Volume"

' Autofit the newly created columns
ws.Columns("I:L").AutoFit
        
' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loops through list
    For i = 2 To LastRow

        'Check if same tickername if not, else
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

              'Set the ticker_name
               ticker_name = ws.Cells(i, 1).Value

               'Add to the ticker_total
               ticker_total = ticker_total + ws.Cells(i, 7).Value

                'Print the ticker_name
                ws.Range("I" & Summary_Table_Row).Value = ticker_name

                'Print the total
                ws.Range("L" & Summary_Table_Row).Value = ticker_total
                                    
                
        'Yearly change from opening price (C) at the beginning of a given year to the closing price (F) at the end of that year.
                open_value = ws.Cells(i, 3).Value
                close_value = ws.Cells(i, 6).Value
                year_change = (close_value - open_value)
                'Print the yearly change
                ws.Range("J" & Summary_Table_Row).Value = year_change
                
                'Conditional formatting highlights positive change in green and negative change in red
                If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                ElseIf ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
          
                
        'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
                percent_change = (year_change / open_value)
                ws.Range("K" & Summary_Table_Row).Value = Format(percent_change, "0.00%")

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'Reset the ticker total
                ticker_total = 0
                

            

        'If the cell immediately following a row is the same
                'Add to the ticker total
                ticker_total = ticker_total + ws.Cells(i, 7).Value
    
        End If

    Next i
    



Next ws

End Sub







