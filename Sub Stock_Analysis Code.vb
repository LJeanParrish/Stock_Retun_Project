Sub Stock_Analysis()

     For Each ws In Worksheets
        Dim WorksheetName As String
        
        ' Capture the WorksheetName
                WorksheetName = ws.Name
  
  
'Part 1 - Create headers for "Ticker," 'Yearly Change," "Percent Change," "Total Stock Volume"
            ws.Cells(1, 9).Value = "Ticker"
            
            ws.Cells(1, 10).Value = "Yearly Change"
        
            ws.Cells(1, 11).Value = "Percent Change"
            
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            ws.Cells(2, 14).Value = "Greatest % Increase"
            
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            
            ws.Cells(4, 14).Value = "Greatest Total Volume"
            
            ws.Cells(1, 15).Value = "Ticker"
            
            ws.Cells(1, 16).Value = "Value"
            
            
'Part 2 - Create variables to hold the key stock inforamtion
            Dim Ticker As String
            Dim Yearly_Change As Double
            Dim Percent_Change As Double
            Dim Stock_Volume As Double
            
            Yearly_Change = 0
            Percent_Change = 0
            Stock_Volume = 0

        ' Create the location to hold each stock variable in a summary table
        
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
            
            Dim startrow As Long
            
            startrow = 2
            

'Part 3 - Capture Tickers from column A
                     
                lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
              
               'Create conditional to loop through tickers
            
     For i = 2 To lastrow
                
                    'Check to see if the next row contains same Ticker
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    'Set the Ticker
                    
                        Ticker = ws.Cells(i, 1).Value
                        
                                                
'Part 4 - Calculate the Yearly Change
    
                    'Create variable for open and close price
                    Dim Yr_Open As Double
                    Dim Yr_Close As Double
        
                    'Capture the last closing price before the new ticker
                    Yr_Close = ws.Cells(i, 6).Value
                 
                    
                    'Capture the first open price of each ticker
                    
                     Yr_Open = ws.Cells(startrow, 3).Value
                                     
                    
                    'Calculate Yearly Change by subtracting closing price from open price
                    
                    Yearly_Change = Yr_Close - Yr_Open
                                                
                         
'Part 5 - Calculate the Percent Change
    
                    'Percent Change equals Yearly Change divided by open price
                    
                    If Yr_Open = 0 Then
                    
                    Percent_Change = 0
                    
                    Else
                    
                        Percent_Change = (Yearly_Change / Yr_Open)
                        
                    End If
                        
                             
'Part 6 - Capture initial Stock volume
                     
                     Stock_Volume = Stock_Volume + (ws.Cells(i, 7).Value)
                     
                                               
                            
'Part 7 - Print all values to Summary Table
                                          
            ' Print the Ticker in the Summary Table
                        
                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                        
            ' Print the Yearly_Change in the Summary Table
                        
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                    
            ' Print the Percent Change to the Summary Table
                        
                    ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        
                    'Convert to a percentage
                        
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                                
            ' Print the Stock Volume to the Summary Table
                        
                    ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
                        
                    
                    
                    'Add conditional to format color on Yearly Change
                    
                        If Yearly_Change > 0 Then
                            
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4 ' Green
                             
                        Else
                
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 ' Red
                            
                        End If
                    
                    
            ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
     
                          Yearly_Change = 0
                                             
                          Percent_Change = 0
                          
                          Stock_Volume = 0
                          
                          startrow = i + 1
                          
                          
'Part 8 - Calculate total stock volume
                          
        ' If the cell immediately following a row is the same ticker
                        
            Else

                ' Add to the Stock_Volume need to figure out correct equation

                Stock_Volume = (Stock_Volume + ws.Cells(i, 7).Value)
                                            
                                
        End If
                
    Next i
    
'Extra Credit 1

    'Calculate the MaxVolume for worksheet
        Dim max_vol As Double
        max_vol = 0
        Dim max_tickernm As String
        
    'Create Conditional to loop through volumes on spreadsheet

        For x = 2 To Summary_Table_Row
    
        If ws.Cells(x, 12).Value > max_vol Then
        
         max_vol = ws.Cells(x, 12).Value
         max_tickernm = ws.Cells(x, 9).Value
         ws.Range("P4").Value = max_vol
         ws.Range("O4").Value = max_tickernm
                 
        End If
    
    Next x
 
 
'Extra Credit 2
    'Calculate the Min Percentage Change for worksheet
    Dim min_percent As Double
    min_percent = 0
    Dim min_tickernm As String
    
     'Create Conditional to loop through volumes on spreadsheet
    
        For y = 2 To Summary_Table_Row
        
        If ws.Cells(y, 11).Value < min_percent Then
         min_percent = ws.Cells(y, 11).Value
         min_tickernm = ws.Cells(y, 9).Value
         ws.Range("P3").Value = min_percent
         ws.Range("O3").Value = min_tickernm
         
        ws.Range("P3").NumberFormat = "0.00%"
                             
         End If
    
    Next y
    
'Extra Credit 3
    'Calculate the Max Percentage Change for worksheet
    Dim max_percent As Double
    max_percent = 0
    Dim max_tickernam As String
    
     'Create Conditional to loop through volumes on spreadsheet
    
        For Z = 2 To Summary_Table_Row
        
        If ws.Cells(Z, 11).Value > max_percent Then
         max_percent = ws.Cells(Z, 11).Value
         max_tickernam = ws.Cells(Z, 9).Value
         ws.Range("P2").Value = max_percent
         ws.Range("O2").Value = max_tickernam
         
        ws.Range("P2").NumberFormat = "0.00%"
                             
         End If
    
    Next Z
    
         ' Autofit to display data
            ws.Columns("A:P").AutoFit
            
            
    Next ws
    
    
End Sub

