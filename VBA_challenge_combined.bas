Attribute VB_Name = "Module1"
Sub VBA_challenge_loop()

    ' Defining WS as a worksheet variable
    Dim WS As Worksheet
    
    ' Looping through each worksheet
    For Each WS In Worksheets
    
        ' Activating the worksheet
        WS.Activate
      
        ' Running the macro below on the worksheet
        Call VBA_challenge
    
    ' Moving on to the next worksheet
    Next WS

End Sub

Sub VBA_challenge()

    'Creating column and/or row headers for the output tables
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    ' Setting an initial variable for holding the ticker symbol
    Dim Ticker As String
    
    ' Setting a variable for holding the ticker's number of entries
    Dim Ticker_Total_Entries As Long
    Ticker_Total_Entries = 0
    
    ' Setting a variable for holding the ticker's opening price at the beginning of the year
    Dim Yearly_Opening_Price As Double
    Yearly_Opening_Price = 0
    
    ' Setting a variable for holding the ticker's closing price at the end of the year
    Dim Yearly_Closing_Price As Double
    Yearly_Closing_Price = 0
    
    ' Setting a variable for holding the difference between the ticker's yearly opening and closing prices
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    ' Setting an initial variable for holding the total volume per ticker
    Dim Ticker_Total_Volume As Double
    Ticker_Total_Volume = 0
    
    ' Setting a variable for and counting the number of rows
    Dim No_of_Rows As Long
    No_of_Rows = Range("A1").End(xlDown).Row
    
    ' Keeping track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    ' Looping through all of the ticker symbols
    For I = 2 To No_of_Rows
    
        ' Checking if the next row has the same ticker symbol; if not
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
            ' Setting the ticker name
            Ticker = Cells(I, 1).Value
            
            ' Adding to the total volume
            Ticker_Total_Volume = Ticker_Total_Volume + Cells(I, 7).Value
            
            ' Printing the ticker in the summary table
            Cells(Summary_Table_Row, 9).Value = Ticker
            
            ' Setting the yearly opening price
            ' MsgBox ("Total Ticker Entries: " & Ticker_Total_Entries)
            ' MsgBox ("i Minus Total Ticker Entries: " & i - Ticker_Total_Entries)
            Yearly_Opening_Price = Cells(I - Ticker_Total_Entries, 3).Value
            ' MsgBox ("Yearly Opening Price: $" & Yearly_Opening_Price)
            
            ' Setting the yearly closing price
            Yearly_Closing_Price = Cells(I, 6).Value
            ' MsgBox ("Yearly Closing Price: $" & Yearly_Closing_Price)
            
            ' Calculating and printing the yearly change in the summary table
            Yearly_Change = Yearly_Closing_Price - Yearly_Opening_Price
            Cells(Summary_Table_Row, 10).Value = Yearly_Change
            
            ' Calculating and printing the percent change in the summary table, if the change isn't 0
            If Yearly_Change <> 0 Then
            
                Cells(Summary_Table_Row, 11).Value = Round((100 * (Yearly_Change / Yearly_Opening_Price)), 2) & "%"
                
            ' If the yearly change is 0
            Else
                
                Cells(Summary_Table_Row, 11).Value = "0%"
                
            End If
            
            ' Printing the total volume in the summary table
            Cells(Summary_Table_Row, 12).Value = Ticker_Total_Volume
            
            ' Adding one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Resetting the total volume, yearly opening price, and yearly closing price
            Ticker_Total_Volume = 0
            Yearly_Opening_Price = 0
            Yearly_Closing_Price = 0
            Ticker_Total_Entries = 0
       
        ' Checking if the next row has the same ticker symbol; if yes
        Else
        
            ' Adding to the total volume
            Ticker_Total_Volume = Ticker_Total_Volume + Cells(I, 7).Value
            
            ' Adding one to the number of ticker entries
            Ticker_Total_Entries = Ticker_Total_Entries + 1
            
        End If
        
    Next I
    
    ' Setting a variable for and counting the number of rows in the summary table
    Dim No_of_ST_Rows As Double
    No_of_ST_Rows = Range("I1").End(xlDown).Row
    
    ' Looping through all of the yearly change values in the summary table
    For i2 = 2 To No_of_ST_Rows
    
        ' Checking if the yearly change is positive; if yes
        If Cells(i2, 10).Value > 0 Then
        
            ' Changing the background color of the yearly change and percent change cells to green
            Cells(i2, 10).Interior.ColorIndex = 4
            Cells(i2, 11).Interior.ColorIndex = 4
            
        ' If the yearly change is negative
        ElseIf Cells(i2, 10).Value < 0 Then
        
            ' Changing the background color of the yearly change and percent change cells to red
            Cells(i2, 10).Interior.ColorIndex = 3
            Cells(i2, 11).Interior.ColorIndex = 3
            
        End If
        
    Next i2
    
    ' Setting initial variables and values for summary table 2
    Dim Greatest_Increase_Ticker As String
    Greatest_Increase_Ticker = Cells(2, 9).Value
    Dim Greatest_Increase_Value As Double
    Greatest_Increase_Value = Cells(2, 11).Value
    Dim Greatest_Decrease_Ticker As String
    Greatest_Decrease_Ticker = Cells(2, 9).Value
    Dim Greatest_Decrease_Value As Double
    Greatest_Decrease_Value = Cells(2, 11).Value
    Dim Greatest_Total_Ticker As String
    Greatest_Total_Ticker = Cells(2, 9).Value
    Dim Greatest_Total_Value As Double
    Greatest_Total_Value = Cells(2, 12).Value
    
    ' Looping through all of the percent change values and total stock volume in the summary table
    For i3 = 2 To No_of_ST_Rows
    
        ' Checking if the next row's percent change is larger; if yes
        If Cells(i3, 11).Value > Greatest_Increase_Value Then
        
            ' Setting a new value for the greatest percent increase ticker and value
            Greatest_Increase_Ticker = Cells(i3, 9).Value
            Greatest_Increase_Value = Cells(i3, 11).Value
            
        End If
        
        ' Checking if the next row's percent change is smaller; if yes
        If Cells(i3, 11).Value < Greatest_Decrease_Value Then
        
            ' Setting a new value for the greatest percent increase ticker and value
            Greatest_Decrease_Ticker = Cells(i3, 9).Value
            Greatest_Decrease_Value = Cells(i3, 11).Value
            
        End If
        
        ' Checking if the next row's total stock volume is larger; if yes
        If Cells(i3, 12).Value > Greatest_Total_Value Then
        
            ' Setting a new value for the greatest percent increase ticker and value
            Greatest_Total_Ticker = Cells(i3, 9).Value
            Greatest_Total_Value = Cells(i3, 12).Value
            
        End If
        
    Next i3
    
    ' Placing all data into the second summary table
    Cells(2, 16).Value = Greatest_Increase_Ticker
    Cells(2, 17).Value = Round(100 * Greatest_Increase_Value, 2) & "%"
    Cells(3, 16).Value = Greatest_Decrease_Ticker
    Cells(3, 17).Value = Round(100 * Greatest_Decrease_Value, 2) & "%"
    Cells(4, 16).Value = Greatest_Total_Ticker
    Cells(4, 17).Value = Greatest_Total_Value
    
End Sub
