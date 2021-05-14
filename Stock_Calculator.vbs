Attribute VB_Name = "Module1"
Sub Stock_Calculator():

'variable to count worksheets
Dim Tabs As Integer

'count worksheets
Tabs = Application.Worksheets.Count

'loop through each worksheet
For j = 1 To Tabs

Worksheets(j).Activate

    'Set variables
    'Stock Ticker
    Dim Stock As String

    'Determine Last row
    LastRow = Cells.SpecialCells(xlCellTypeLastCell).Row

    'Stock volume variable
    Dim Stock_Volume As Double
    Stock_Volume = 0

    'varibale to keep track of stock ticker location
    Dim Stock_Row_Location As Integer
    Stock_Row_Location = 2
    
    'set variables of opening price
    Dim Opening_price As Double
    Opening_price = Cells(2, 3).Value
    
    'bold headers
    Range("J1:M1").Font.Bold = True
    'adjust column width
    
    Columns("J:M").ColumnWidth = 15
    'format yearly change into currency
    Columns("K").NumberFormat = "$#,##0.00"

    'format percent change into percentage
    Columns("L").NumberFormat = "#.###%"
    
    'set variables of closing price
    Dim Closing_price As Double

    'insert headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Volume"

    'Loop through all Stock tickers
    For i = 2 To LastRow


        'Check if same stock ticker symbol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            'Set the stock ticker symbol
            Stock = Cells(i, 1).Value
        
            'Add volume
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
            'last row of current stock
            Closing_price = Cells(i, 6).Value
            
            'Print the diffence of closing price to opening price
            Range("K" & Stock_Row_Location).Value = Closing_price - Opening_price
            
                    If Closing_price - Opening_price > 0 Then
                    'make green if positive
                    Range("K" & Stock_Row_Location).Interior.ColorIndex = 4
                    
                    ElseIf Closing_price - Opening_price < 0 Then
                    'make red if negative
                    Range("K" & Stock_Row_Location).Interior.ColorIndex = 3
                    
                    Else
                    'make red if negative
                    Range("K" & Stock_Row_Location).Interior.ColorIndex = 2
                    
                    'End loop
                    End If
                    
                
                'check if opening price is 0
                If Opening_price = 0 Then
                
                'set opening_price to 0
                Range("L" & Stock_Row_Location).Value = 0
                
                Else
                
                'Print the percent changeof closing price to opening price
                Range("L" & Stock_Row_Location).Value = (Closing_price - Opening_price) / Opening_price
                End If
                
            'Print stock ticker symbol
            Range("J" & Stock_Row_Location).Value = Stock
        
            'Print the sum of stock volume
            Range("M" & Stock_Row_Location).Value = Stock_Volume
        
            
            'Add next ticker symbol
            Stock_Row_Location = Stock_Row_Location + 1
            
            'first row of next stock
            Opening_price = Cells(i + 1, 3).Value
            
            'Reset volume
            Stock_Volume = 0
            
        Else
     
        'Add volume
        Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
        'end stock loop
        End If
        
    'increment for loop
    Next i

'increment worksheet
Next j

'end sub-routine

End Sub
