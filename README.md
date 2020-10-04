# VBA-Homework-Zhiheng

Sub stock_market()

  ' Set an initial variable for holding the ticker name
  Dim ticker_Name As String

  ' Set an initial variable for holding the total volume per ticker
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Create columns and cell names
  Cells(1, 10).Value = "Ticker"
  Cells(1, 11).Value = "Total Stock Volume"
  Cells(1, 12).Value = "Percentage Change"
  Cells(1, 13).Value = "Yearly Change"
  Cells(2, 15).Value = "Greatest % Increase"
  Cells(3, 15).Value = "Greatest % Decrease"
  Cells(4, 15).Value = "Greatest Total Volume"
  Cells(1, 16).Value = "Ticker"
  Cells(1, 17).Value = "Value"
  
  ' Set initial variable for holding the open price and close price of each stock
  ' Also set the initial variable for holding the percentage change of stock return
  Dim Open_price As Double
  Dim Close_price As Double
  Dim Percentage_Change As Double
  
  
  Dim Ticker_lookup1 As String
  Dim Ticker_lookup2 As String
  Dim Ticker_lookup3 As String
  Dim Yearly_Change As Double
  Dim Greatest_Increase As Double
  Dim Greatest_decrease As Double
  Dim Greatest_volume As Double
  
  Greatest_volume = 0


  'Set the initial value of open price
  Open_price = Cells(2, 3).Value

  ' Loop through all stock trading transactions
    For i = 2 To 760192

    ' Check if we are still within the same stock, if it is not...
          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
      ' Set the Ticker name and find its close price
          ticker_Name = Cells(i, 1).Value
      
          Close_price = Cells(i, 6).Value

      ' Add to the Stock Total Volume
          Volume_Total = Volume_Total + Cells(i, 7).Value
      
      'Calculate the percentage change of stock price
          If Open_price <> 0 Then
                Percentage_Change = Round(((Close_price - Open_price) / Open_price) * 100, 2)
          Else
          End If
          
      'Calculate the yearly change of stock price and highlight them
          Yearly_Change = Close_price - Open_price
           If Yearly_Change > 0 Then
                Range("M" & Summary_Table_Row).Interior.ColorIndex = 4
           Else
                Range("M" & Summary_Table_Row).Interior.ColorIndex = 3
           End If
      
      ' Print the Ticker in the Summary Table
         Range("J" & Summary_Table_Row).Value = ticker_Name

      ' Print the Volume total, annual Stock return to the Summary Table
         Range("K" & Summary_Table_Row).Value = Volume_Total
      
         Range("L" & Summary_Table_Row).Value = Percentage_Change & "%"
            
         Range("M" & Summary_Table_Row).Value = Yearly_Change
    
      ' Set a condition where the highest stock return should equal to the highest price percentage change, if it is less than price percentage change, then:
    If Greatest_Increase < Percentage_Change Then
        Greatest_Increase = Percentage_Change
        Ticker_lookup1 = Cells(i, 1).Value
    End If
    
        
      ' Set a condition where the decreasing stock return should equal to the decreasing price percentage change, if it is larger than price percentage change, then:
    If Greatest_decrease > Percentage_Change Then
        Greatest_decrease = Percentage_Change
        Ticker_lookup2 = Cells(i, 1).Value
    End If
      
      ' Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total and open price for next stock
         Volume_Total = 0
         Open_price = Cells(i + 1, 3).Value

    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value
  
    If Volume_Total > Greatest_volume Then
        Greatest_volume = Volume_Total
        Ticker_lookup3 = Cells(i, 1).Value
    End If
     
      
    End If

  Next i

' print out the greatest stock return increase and decrease as well as corresponding tickers
Range("Q2").Value = Greatest_Increase & "%"
Range("Q3").Value = Greatest_decrease & "%"
Range("Q4").Value = Greatest_volume
Range("P2").Value = Ticker_lookup1
Range("P3").Value = Ticker_lookup2
Range("P4").Value = Ticker_lookup3
End Sub


