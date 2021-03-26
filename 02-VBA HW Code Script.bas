Attribute VB_Name = "Module1"
Sub VPStockMarketAnalysis_Final()

'-----------------------------------------------------------------------------------------------
' SUMMARIZE YEARLY CHANGE, PERCENTAGE CHANGE & VOLUME BY TICKER FOR EACH WORKSHEET
'-----------------------------------------------------------------------------------------------
    'Declare ws as Worksheet
    
    Dim ws As Worksheet
    
    'Set for loop to perform all actions on each worksheet
    
    For Each ws In Worksheets
    
'------------------------------------------------------------------------------------------------
    
    'Create Column Headers for Results
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Bold Header Text
    
    ws.Range("I1:L1").Font.Bold = True
    
'-----------------------------------------------------------------------------------------------
' Place Ticker in column I
'-----------------------------------------------------------------------------------------------
      
    'Establish last row of each worksheet
    
    Dim LastRow As Long
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Declare & Assign variable to hold ticker
    
    Dim TickerSymbol As String
    
    'Create counter for placement of ticker in colum I
    
    Dim Counter As Integer
    
    Counter = 2
    
'-----------------------------------------------------------------------------------------------
    
    
    'Create loop to populate ticker in column I
    
    For Tickercounter = 2 To LastRow
    
        'Create conditional to search row by row until ticker changes
    
        If ws.Cells(Tickercounter, 1).Value <> ws.Cells(Tickercounter - 1, 1) Then
    
            'Store ticker in TickerSymbol
    
            TickerSymbol = ws.Cells(Tickercounter, 1).Value
    
            'Place ticker in column I
    
            ws.Cells(Counter, 9) = TickerSymbol
    
                'Move ticker placement reference to next row
                Counter = (Counter + 1)
    
        End If
    
    Next Tickercounter
    
    
'-----------------------------------------------------------------------------------------------
' Calculate Yearly Change and Percentage Change
'-----------------------------------------------------------------------------------------------
    
    'Setup variables to use in calculation
    
    Dim OpenPrice As Double
    
    OpenPrice = ws.Cells(2, 3)
    
    
    Dim ClosePrice As Double
    
    Dim YearChange As Double
    
    Dim PerChg As Double
    
    
    Dim YrChgCounter As Integer
    
    YrChgCounter = 2
    
'-----------------------------------------------------------------------------------------------
    
    'Create loop to calculate yearly change and percentage change for each ticker
        
    For Pricecounter = 2 To LastRow
    
        'Look for change in Ticker to calculate Yearly Change
    
        If ws.Cells(Pricecounter, 1).Value <> ws.Cells(Pricecounter + 1, 1).Value Then
    
            ClosePrice = ws.Cells(Pricecounter, 6).Value
    
        
                YearChange = ClosePrice - OpenPrice
    
                    ws.Cells(YrChgCounter, 10).Value = YearChange
    
            'Calculate Percentage Change
            
            If ClosePrice = 0 Or OpenPrice = 0 Then
            
                ws.Cells(YrChgCounter, 11).Value = 0
    
            ElseIf ClosePrice <> 0 And OpenPrice <> 0 Then
            
                 PerChg = YearChange / OpenPrice
    
                    ws.Cells(YrChgCounter, 11).Value = PerChg
    
                        'Change interior PerChg on cells based on increase/decrease in yearl change
    
                        If YearChange > 0 Then
    
                            ws.Cells(YrChgCounter, 10).Interior.ColorIndex = 43
    
                        Else: ws.Cells(YrChgCounter, 10).Interior.ColorIndex = 3
    
                        End If
    
            End If
    
            'Re-set counters for next ticker calculation
            
            YrChgCounter = YrChgCounter + 1
        
            OpenPrice = ws.Cells(Pricecounter + 1, 3)
        
            'Format Percent Change as pecent
            
            ws.Columns("K:K").NumberFormat = "0.00%"
        
    
        End If
    
   
    Next Pricecounter
    
'-----------------------------------------------------------------------------------------------
' Calculate sum of volume for each ticker
'-----------------------------------------------------------------------------------------------
    
    'Setup variables for calculation of volume sum
    
    Dim VolumeSum As Double
    
    VolumeSum = 0
    
    VolCounter = 2
    
        'Create loop to capture total volume
    
        For Volume = 2 To LastRow
    
            If ws.Cells(Volume, 1).Value = ws.Cells(Volume + 1, 1).Value Then
    
            VolumeSum = VolumeSum + ws.Cells(Volume, 7).Value
    
        
            ElseIf ws.Cells(Volume, 1).Value <> ws.Cells(Volume + 1, 1).Value Then
    
                VolumeSum = VolumeSum + ws.Cells(Volume, 7).Value
    
                    ws.Cells(VolCounter, 12) = VolumeSum
    
                        VolCounter = VolCounter + 1
    
                        VolumeSum = 0
    
            End If
    
        Next Volume
                           
            
'-----------------------------------------------------------------------------------------------
' Calculate Greatest Perentage Increase & Total Volume, Greatest Percentage Decrease
'-----------------------------------------------------------------------------------------------
    ' Create variables to calculate Greatest Percentage Increase
    
    Dim Greatest As Double
  
    Dim Total As Double
    
    Dim Grand_Total As Double
    
    Grand_Total = 0
    
    Dim Total_Hold As Double
    
    Dim Ticker_Hold As String
    
    Dim Grand_Ticker As String
    
'------------------------------------------------------------------------------------------------
            
            'Create loop to calcuate greatest percentage increase
            
            For Total_Row = 2 To LastRow
        
                If ws.Cells(Total_Row + 1, 11).Value > ws.Cells(Total_Row, 11).Value Then
                
                    Total = ws.Cells(Total_Row + 1, 11).Value
                    
                    Ticker_Hold = ws.Cells(Total_Row + 1, 9).Value
                
                ElseIf ws.Cells(Total_Row + 1, 11).Value < ws.Cells(Total_Row, 11).Value Then
    
                    Total = ws.Cells(Total_Row, 11).Value
                    
                    Ticker_Hold = ws.Cells(Total_Row, 9).Value
                    
                End If
                
                If Total > Grand_Total Then
                
                    Grand_Total = Total
                                       
                                       
                ElseIf Total < Grand_Total Then
                
                    Grand_Total = Grand_Total
                    
                                  
                End If
                  
                
            Next Total_Row
            
                For Grand_Ticker_Row = 2 To LastRow
                        
                        If ws.Cells(Grand_Ticker_Row, 11) = Grand_Total Then
                            
                            Grand_Ticker = ws.Cells(Grand_Ticker_Row, 9)
                        
                        End If
                    
                    Next Grand_Ticker_Row
                
                                        
                    ws.Cells(2, 15) = "Greatest % Increase"
                    
                    ws.Cells(2, 16) = Grand_Ticker
                    
                    ws.Cells(2, 17) = Grand_Total
                    
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                
  ' Create variables to calculate Greatest Percentage Decrease
    
    Dim Least As Double
  
    Dim Total_D As Double
    
    Dim Grand_Total_D As Double
    
    Grand_Total_D = 0
    
    Dim Total_Hold_D As Double
    
    Dim Ticker_Hold_D As String
    
    Dim Grand_Ticker_D As String
    
'------------------------------------------------------------------------------------------------
                 
              
  'Create loop to calcuate greatest percentage decrease
            
            For Total_Row_D = 2 To LastRow
        
                If ws.Cells(Total_Row_D + 1, 11).Value < ws.Cells(Total_Row_D, 11).Value Then
                
                    Total_D = ws.Cells(Total_Row_D + 1, 11).Value
                    
                    Ticker_Hold_D = ws.Cells(Total_Row_D + 1, 9).Value
                
                ElseIf ws.Cells(Total_Row_D + 1, 11).Value > ws.Cells(Total_Row_D, 11).Value Then
    
                    Total_D = ws.Cells(Total_Row_D, 11).Value
                    
                    Ticker_Hold_D = ws.Cells(Total_Row_D, 9).Value
                    
                End If
                
                If Total_D > Grand_Total_D Then
                
                    Grand_Total_D = Grand_Total_D
                                       
                                       
                ElseIf Total_D < Grand_Total_D Then
                
                    Grand_Total_D = Total_D
                    
                                  
                End If
                  
                
            Next Total_Row_D
            
                For Grand_Ticker_Row_D = 2 To LastRow
                        
                        If ws.Cells(Grand_Ticker_Row_D, 11) = Grand_Total_D Then
                            
                            Grand_Ticker_D = ws.Cells(Grand_Ticker_Row_D, 9)
                        
                        End If
                    
                    Next Grand_Ticker_Row_D
                
                                        
                    ws.Cells(3, 15) = "Greatest % Decrease"
                    
                    ws.Cells(3, 16) = Grand_Ticker_D
                    
                    ws.Cells(3, 17) = Grand_Total_D
                    
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                    
                    
                    
              ' Create variables to calculate Greatest Volume Increase
    
    Dim Vol_Increase As Double
  
    Dim Total_V As Double
    
    Dim Grand_Total_V As Double
    
    Grand_Total_V = 0
    
    Dim Total_Hold_V As Double
    
    Dim Ticker_Hold_V As String
    
    Dim Grand_Ticker_V As String
    
'------------------------------------------------------------------------------------------------
            
            'Create loop to calcuate greatest volume increase
            
            For Total_Row_V = 2 To LastRow
        
                If ws.Cells(Total_Row_V + 1, 12).Value > ws.Cells(Total_Row_V, 12).Value Then
                
                    Total_V = ws.Cells(Total_Row_V + 1, 12).Value
                    
                    Ticker_Hold_V = ws.Cells(Total_Row_V + 1, 9).Value
                
                ElseIf ws.Cells(Total_Row_V + 1, 12).Value < ws.Cells(Total_Row_V, 12).Value Then
    
                    Total_V = ws.Cells(Total_Row_V, 12).Value
                    
                    Ticker_Hold_V = ws.Cells(Total_Row_V, 9).Value
                    
                End If
                
                If Total_V > Grand_Total_V Then
                
                    Grand_Total_V = Total_V
                                       
                                       
                ElseIf Total_V < Grand_Total_V Then
                
                    Grand_Total_V = Grand_Total_V
                    
                                  
                End If
                  
                
            Next Total_Row_V
            
                For Grand_Ticker_Row_V = 2 To LastRow
                        
                        If ws.Cells(Grand_Ticker_Row_V, 12) = Grand_Total_V Then
                            
                            Grand_Ticker_V = ws.Cells(Grand_Ticker_Row_V, 9)
                        
                        End If
                    
                    Next Grand_Ticker_Row_V
                
                                        
                    ws.Cells(4, 15) = "Greatest % Increase"
                    
                    ws.Cells(4, 16) = Grand_Ticker_V
                    
                    ws.Cells(4, 17) = Grand_Total_V
                    
                                    
            'Widen columns to fit data
            
            ws.Columns("I:Q").AutoFit
       
    Next ws

End Sub







