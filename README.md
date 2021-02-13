# VBA-Challenge

Sub Stock_Volume()
   
    ' LOOP THROUGH ALL SHEETS
    
    For Each ws In Worksheets
  
        ' Created a Variable to hold Ticker
        Dim Ticker As String
        
        ' Set an initial variable for holding the total volume per ticker
        Dim Volume_Total As Double
        Volume_Total = 0
        
        ' Add the headers to Columns I and J
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        
        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Keep track of the location for Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
    
            ' Loop through all Tickers trading
            For i = 2 To lastRow
        
                ' Check if we are still within the same Ticker, if it is not
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                  ' Set the Ticker name
                  Ticker = ws.Cells(i, 1).Value
            
                  ' Add to the Volume Total
                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
                  ' Print the Ticker in the Summary Table
                  ws.Range("I" & Summary_Table_Row).Value = Ticker
            
                  ' Print the Trading Volume to the Summary Table
                  ws.Range("J" & Summary_Table_Row).Value = Volume_Total
            
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
                  
                  ' Reset the Volume Total
                  Volume_Total = 0
            
                ' If the cell immediately following a row is the same Ticker
                Else
            
                  ' Add to the Volume Total
                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
                End If
        
          Next i
   
   
        ' Add Column and header for Yearly Change
            ws.Range("J1").EntireColumn.Insert
            ws.Cells(1, 10).Value = "Yearly Change"
                       
        ' Add Column and header for Percent Change
            ws.Range("K1").EntireColumn.Insert
            ws.Cells(1, 11).Value = "Percent Change"
   
   
        ' Locate the start and end row of each ticker to identify where to pull open and close prices from
        Dim TickerStartRow As Double
        Dim TickerEndRow As Double

        ' Keep track of the Open, Close and Price Change in the summary table
        Dim Open_Price_Row As Double
        Dim Closing_Price_Row As Double
        Dim Open_Price As Double
        Dim Closing_Price As Double
        Dim Price_Change As Double
        Dim Price_Change_P As Double
        Dim Summary_Table_Row_2 As Double
        Summary_Table_Row_2 = 2
        
        ' Set Last row for summary table
        lastRow_S = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Loop through all Tickers trading to find Open & Close Prices for the year
            For i = 2 To lastRow_S
            
                
                Open_Price_Row = WorksheetFunction.Match(ws.Cells(i, 9).Value, ws.Range("A:A"), 0)
                Open_Price = ws.Cells(Open_Price_Row, 3).Value

                Closing_Price_Row = Open_Price_Row + WorksheetFunction.CountIf(ws.Range("A:A"), ws.Cells(i, 9).Value) - 1
                Closing_Price = ws.Cells(Closing_Price_Row, 6).Value

                ' Calculate the Price Change
                Price_Change = Closing_Price - Open_Price

                ' Print the Price Change to the Summary Table
                ws.Range("J" & Summary_Table_Row_2).Value = Price_Change

                ' Calculate the Price Change %
                If Open_Price = 0 Then
                    Open_Price = 0.01
                End If

                Price_Change_P = Price_Change / Open_Price

                ' Print the Price Change % to the Summary Table
                ws.Range("K" & Summary_Table_Row_2).Value = Price_Change_P

                ' Add one to the summary table row
                Summary_Table_Row_2 = Summary_Table_Row_2 + 1

                ' Reset the all Price variables
                Open_Price_Row = 0
                Closing_Price_Row = 0
                Open_Price = 0
                Closing_Price = 0
                Price_Change = 0
                Price_Change_P = 0


            Next i

    
    ' Conditional Formatting to highlight cells in summary table: positive change = green (4) and negative change = red (3)
        For i = 2 To lastRow_S
        
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If


            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
                ws.Cells(i, 11).NumberFormat = "0.00%"
            Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
                ws.Cells(i, 11).NumberFormat = "0.00%"
            End If

        Next i
                
    ' Add the headers for the second summary table
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"


    ' Create another summary table with Ticker & Value columns for Greatest % increase / decrease and greatest total volume
    Dim Max_Increase As Double
    Dim Max_Decrease As Double
    Dim Max_Volume As Double

    Max_Increase = WorksheetFunction.Max(ws.Range("K2:K" & lastRow_S))
    Max_Decrease = WorksheetFunction.Min(ws.Range("K2:K" & lastRow_S))
    Max_Volume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow_S))

    'Print Max and Min results to table
    ws.Range("P2").Value = Max_Increase
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").Value = Max_Decrease
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("P4").Value = Max_Volume

   

                                                              
    ' Loop through Summary Table
      For i = 2 To lastRow_S

          ' Check if Max Increase matches Summary table
          If Max_Increase = ws.Cells(i, 11).Value Then

              ' Retrieve the Ticker name
              ws.Range("O2").Value = ws.Cells(i, 9).Value

          ' Check if Max Decrease matches Summary table
          ElseIf Max_Decrease = ws.Cells(i, 11).Value Then

              ' Retrieve the Ticker name
              ws.Range("O3").Value = ws.Cells(i, 9).Value

           ' Check if Max Volume matches Summary table
          ElseIf Max_Volume = ws.Cells(i, 12).Value Then

              ' Retrieve the Ticker name
              ws.Range("O4").Value = ws.Cells(i, 9).Value

          End If

     Next i

   
   
    Next ws

End Sub
