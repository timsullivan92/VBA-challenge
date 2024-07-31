Attribute VB_Name = "Module1"
Sub Stockmetrics()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

'Variable to hold column name
Dim column As Integer
column = 1

'Variable to increment the Ticker symbols in column I
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Variable to Sum Stock Volume for each ticker
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Variables for calculating Quarterly Change/Percent Change
Dim Open_Price As Double
Dim End_Price As Double
Dim Quarterly_Change As Double
Dim Percent_Change As Double

'Variables for calculating Greatest Increase %, Greatest Decrease %, Greatest Volume
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double
Dim GI_row As Integer
Dim GD_row As Integer
Dim GV_row As Integer


'Loop Counter
Dim loop_count As Integer
counter = 0

Dim i As Long
Dim Ticker_Symbol As String


'Find Last Row
Dim lastRow As Long
lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row

'Create Column Headers for Ticker,etc.
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Create Column Headers for Greatest Increase, Decrease, etc.
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Format Columns J,K,L
ws.Range("J:J").ColumnWidth = 20
ws.Range("K:K").ColumnWidth = 20
ws.Range("L:L").ColumnWidth = 20
ws.Range("K:K").NumberFormat = "0.00%"

'Format Columns O,P,Q

ws.Range("Q2,Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "0"
ws.Range("O:O").ColumnWidth = 20
ws.Range("Q:Q").ColumnWidth = 20


'Initialize Open Price
Open_Price = ws.Cells(2, 3).Value
'ws.Range("J" & Summary_Table_Row).Value = Open_Price

'Loop through rows in the column
For i = 2 To lastRow
'Searches for when the value of the next cell is different than the current cell
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
    
'Write Ticker Symbol to Column I
        Ticker_Symbol = Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
        
'Write Total Stock Volume to Column L
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
'Calculate Open Price
        Open_Price = ws.Cells(i - loop_count, 3).Value
        
'Calculate End_Price
        End_Price = ws.Cells(i, 6).Value
             
'Write Quarterly Change to Colum J
        Quarterly_Change = End_Price - Open_Price
        
        ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
        
'Format Quarterly Change fill color
                If (Quarterly_Change > 0) Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
                ElseIf (Quarterly_Change < 0) Then
                    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            
                End If
        
        
'Write Percent Change to Column K
        Percent_Change = Quarterly_Change / Open_Price
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
'Increment Summary_Table_Row
        
        Summary_Table_Row = Summary_Table_Row + 1
        
'Reset loop_count to zero
        loop_count = 0
'Reset Total_Stock_Volume to zero
        Total_Stock_Volume = 0
    
    Else
'Calculate Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
'Increment loop_count variable
        loop_count = loop_count + 1
   
    End If
Next i

'Find last row of new columns
lastRow_newcol = ws.Range("I" & Rows.Count).End(xlUp).Row

'Find Greatest % Increase, Greatest % Decrease, Greatest Total Volume
  
    Greatest_Increase = WorksheetFunction.Max(ws.Range("K2:K" & lastRow_newcol))
    ws.Cells(2, 17).Value = Greatest_Increase
    GI_row = Application.WorksheetFunction.Match(Greatest_Increase, ws.Range("K2:K" & lastRow_newcol), 0) + 1
    ws.Cells(2, 16).Value = ws.Cells(GI_row, "I").Value
    
    Greatest_Decrease = WorksheetFunction.Min(ws.Range("K2:K" & lastRow_newcol))
    ws.Cells(3, 17).Value = Greatest_Decrease
    GD_row = Application.WorksheetFunction.Match(Greatest_Decrease, ws.Range("K2:K" & lastRow_newcol), 0) + 1
    ws.Cells(3, 16).Value = ws.Cells(GD_row, "I").Value
    
    Greatest_Volume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow_newcol))
    ws.Cells(4, 17).Value = Greatest_Volume
    GV_row = Application.WorksheetFunction.Match(Greatest_Volume, ws.Range("L2:L" & lastRow_newcol), 0) + 1
    ws.Cells(4, 16).Value = ws.Cells(GV_row, "I").Value
    
Next ws

End Sub
