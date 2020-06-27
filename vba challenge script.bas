Attribute VB_Name = "Module1"
Sub Wall_street()

'Set Ticker Variable
    Dim ticker_Name As String
    
'Declare variables for old and new
    Dim open_date As Double
    open_date = Cells(2, 3).Value
    Dim close_date As Double
    

'Set total stock volume variable and total
    Dim stock_volume As Double
    stock_volume = 0

'Keep track of the location for ticker in summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
'Determine last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop though stock volume
    For i = 2 To LastRow

'Check if we are still within the same Ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
'Set the close date
close_date = Cells(i, 6).Value

'Calculate the yearly change
Dim year_change As Double
year_change = close_date - open_date

'Set the ticker symbol
    ticker_Name = Cells(i, 1).Value

'Set the Total Stock volume
    stock_volume = stock_volume + Cells(i, 7)

'Print the ticker symbol in summary table
    Range("I" & Summary_Table_Row).Value = ticker_Name

'Print the total Stock volume in summary table
    Range("L" & Summary_Table_Row).Value = stock_volume

'Print the yearly in summary table
   Range("J" & Summary_Table_Row).Value = year_change

'Set an if/then for if year change is 0

    If open_date = 0 Then
    Range("K" & Summary_Table_Row).Value = "%" & (0)

Else

'Print the Percent Change in Summary table
    Range("K" & Summary_Table_Row).Value = "%" & (year_change / open_date) * 100
    
End If

If open_date = 0 Then
Range("K" & Summary_Table_Row).Value = "%" & (0)

End If


'Setting open date for next ticker

 open_date = Cells(i + 1, 3).Value
 
 'Add range for positive conditional formating
If year_change < 0 Then
Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

Else
Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

End If

'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1

'Reset total stock volume
    stock_volume = 0
    

'If the cell immediately following a row is the same ticker symbol
    Else

    stock_volume = stock_volume + Cells(i, 7).Value

End If

    Next i

End Sub

