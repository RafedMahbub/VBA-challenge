Attribute VB_Name = "Module1"

Sub StockData()
    
    For Each ws In Worksheets
    
        ' Create Variables to Hold Ticker name, Start date and End date, Open and Closing Price, Yearly change, Stock volume
        Dim Ticker_Name As String
        Dim Start_Date As Long
        Dim End_Date As Long
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Percentage_Change As Double
        Dim Stock_Volume As Double
        Stock_Volume = 0
        Dim Perc_Increase As Double
        Perc_Increase = 0
        Dim Perc_Decrease As Double
        Perc_Decrease = 0
        Dim Gr_Total_Volume As Double
        Gr_Total_Volume = 0
        
        
        ' Keep track of the location for each Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
              
        ' Sort Ticker data
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("A1"), Order:=xlAscending
            .SortFields.Add Key:=ws.Range("B1"), Order:=xlAscending
            .SetRange ws.Range("A1", ws.Range("G1").End(xlDown))
            .Header = xlYes
            .Apply
        End With
        
        
        
         ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
               
        ' Add the word Ticker to the Ninth Column Header
        ws.Cells(1, 9).Value = "Ticker"
        
        ' Add the word Yearly Change to the Tenth Column Header
        ws.Cells(1, 10).Value = "Yearly_Change"
        
        ' Add the word Percentage Change to the Eleventh Column Header
        ws.Cells(1, 11).Value = "Percentage_Change"
        
        ' Add the word Total stock volume to the Twelfth Column Header
        ws.Cells(1, 12).Value = "Total_Stock_Volume"
        
        'Start data for the first ticker
        Start_Date = ws.Cells(2, 2).Value
        Opening_Price = ws.Cells(2, 3).Value
        
        ' Store values in appropriate column by looping through and naming each
        For i = 2 To LastRow
        
            ' Check if we are still within the same Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the Ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Set the End Date
                End_Date = ws.Cells(i, 2).Value
                Closing_Price = ws.Cells(i, 6).Value
                
                'Calculate yearly change
                Yearly_Change = Opening_Price - Closing_Price
                Percentage_Change = Yearly_Change / Opening_Price
                
                ' Add to Total Stock volume
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
                ' Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                ' Print the Yearly change in the Summary Table with colour green if positive and red if negetive
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                If (Yearly_Change < 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                    
                
                ' Print Percentage change in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ' Print the Total Stock_Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Opening price for the next ticker
                Opening_Price = ws.Cells(i + 1, 3).Value
                
                ' Reset the Total Stock Volume
                Stock_Volume = 0
            
            ' If the cell immediately following a row is the same Ticker...
            Else

                ' Add to Total Stock Volume
                 Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

            End If
            
        Next i
        
        'set Header for Ticker and corrosponding value
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 16).Value = "Ticker"
        
        'set Row description
        ws.Cells(2, 15).Value = "Greatest % Increase"
        
        'Find the Maximum price increase and print it
        Perc_Increase = WorksheetFunction.Max(ws.Range("K2 : K" & Summary_Table_Row))
        ws.Cells(2, 17).Value = Perc_Increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        'Find corrosponding Ticker for Maximum price increse
        ws.Cells(2, 16).Value = WorksheetFunction.Index(ws.Range("I2 : I" & Summary_Table_Row), WorksheetFunction.Match(Perc_Increase, ws.Range("K2 : K" & Summary_Table_Row), 0))
        
        'set Row description
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        
        'Find the Maximum price Decrease and print it
        Perc_Decrease = WorksheetFunction.Min(ws.Range("K2 : K" & Summary_Table_Row))
        ws.Cells(3, 17).Value = Perc_Decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'Find corrosponding Ticker for Maximum price Decrese
        ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I2 : I" & Summary_Table_Row), WorksheetFunction.Match(Perc_Decrease, ws.Range("K2 : K" & Summary_Table_Row), 0))
        
        'set Row description
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Find the Greates total volume and print it
        Gr_Total_Volume = WorksheetFunction.Max(ws.Range("L2 : L" & Summary_Table_Row))
        ws.Cells(4, 17).Value = Gr_Total_Volume
        
        'Find corrosponding Ticker for Greates total volume
        ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I2 : I" & Summary_Table_Row), WorksheetFunction.Match(Gr_Total_Volume, ws.Range("L2 : L" & Summary_Table_Row), 0))
        
    
    Next ws
    MsgBox ("End of sheet")


End Sub
