Attribute VB_Name = "Module1"
Sub wallstreet()

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per ticker
  Dim Ticker_Total As Double
  Ticker_Total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
  Dim flag As Integer
    flag = 0
  
  Dim yearly_change As Double
    yearly_change = 0
  
  Dim percent_change As Double
    percent_change = 0
  
  Dim first As Double
    first = 0
  
  Dim last As Double
    last = 0
  
  Dim Biggestincrease As Double
    Biggestincrease = 0
  
  Dim Biggestdecrease As Double
    Biggestdecrease = 0
  
  Dim Biggestvolume As Double
    Biggestvolume = 0

' Loop through all stock purchases
    For i = 2 To 70926

' Check if we are still within the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the ticker name
    Ticker_Name = Cells(i, 1).Value

' Add to the ticker Total
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
'Label all the headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("Percent_Change").NumberFormat = "0.00%"

' Print the ticker name in the Summary Table
    Range("I" & Summary_Table_Row).Value = Ticker_Name

' Print the ticker name to the Summary Table
    Range("L" & Summary_Table_Row).Value = Ticker_Total

    If Ticker_Total > Biggestvolume Then
    
'print the results
    Biggestvolume = Ticker_Total
    Range("Q4").Value = Biggestvolume
    Range("P4").Value = Ticker_Name
    
        Else
    End If
       
' Reset the ticker Total
    Ticker_Total = 0
      
    last = Cells(i, 6).Value
    
    yearly_change = last - first
      
    percent_change = yearly_change / first
    
    If (percent_change > Biggestincrease) Then
    
'print the results
    Biggestincrease = percent_change
    Range("Q2").Value = Format(Biggestincrease, ".00%")
    Range("P2").Value = Ticker_Name

        Else
    End If
    
      If (percent_change < Biggestdecrease) Then
    
'print the results
    Biggestdecrease = percent_change
    Range("Q3").Value = Format(Biggestdecrease, ".00%")
    Range("P3").Value = Ticker_Name

        Else
    End If
      
'print the results of your calculations
     
    Range("J" & Summary_Table_Row).Value = yearly_change
      
    Range("K" & Summary_Table_Row).Value = percent_change
        
'setting conditional formatting to red or green
     
   If yearly_change > 0 Then
    
    Range("J" & Summary_Table_Row).Interior.Color = vbGreen
    
    Else
    
    Range("J" & Summary_Table_Row).Interior.Color = vbRed
    
    End If
    
        
' reset all the declared items change total
      
    yearly_change = 0
      
    percent_change = 0
      
    first = 0
      
    last = 0

' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
      
    flag = 0

' If the cell immediately following a row is the same stock...
        Else

' Add to the stock Total
    Ticker_Total = Ticker_Total + Cells(i, 7).Value
      
    If flag = 0 Then
    
    first = Cells(i, 3).Value
    flag = flag + 1
    
        Else
    End If
    
    End If
Next i


End Sub









