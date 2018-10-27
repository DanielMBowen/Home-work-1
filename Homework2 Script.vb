Sub credit_card()

  ' Set an initial variable for holding the brand name
  Dim Ticker_Name As String
  Dim Year as Integer
  
  
  LastRow = Range("A" & Rows.Count).End(xlUp).Row




  
  Dim Stock_Total As Double
  Stock_Total = 0

  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

 
  For i = 2 To LastRow


   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

     
      Ticker_Name = Cells(i, 1).Value
      Year = Cells(i, 3).Value


    
      Stock_Total = Stock_Total + Cells(i, 7).Value

      

     
      Range("J" & Summary_Table_Row).Value = Ticker_Name

    
      Range("K" & Summary_Table_Row).Value = Year

      Range("L" & Summary_Table_Row).Value = Stock_Total

    Cells(1, 10).Value = "Ticker SymboL"
    Cells(1, 11).Value = "Year"
    Cells(1, 12).Value = "Total Volume"
      Summary_Table_Row = Summary_Table_Row + 1
      
     
      Stock_Total = 0

    Else

      
      Stock_Total = Stock_Total + Cells(i, 7).Value

    End If

  Next i
  

End Sub
