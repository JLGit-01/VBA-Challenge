Sub alpha3()

  ' Set an initial variable for holding the brand name
  Dim Stock_Letter As String

  ' Set an initial variables for stock totals, closinng stock, opening stock
  Dim Stock_Total As Double
  Stock_Total = 0

  Dim Closing_Stock As Double

  Dim Opening_Stock As Double
   

  ' Keep track of the location for each stock detail brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock detail 
  For row = 2 To (ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row)


' Check if we are in the first row of the stock


    If Cells(row - 1, 1).Value <> Cells(row, 1).Value Then
    Opening_Stock = Cells(row,3).Value

    End If

    If Cells(row, 1).Value = Cells(row + 1, 1).Value Then

      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(row, 7).Value

    End if


    If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then

      ' Set the Stock Ticker name
      Stock_Letter = Cells(row, 1).Value
      Closing_Stock = Cells(row,6).Value

      ' Add to the Brand Total
      Stock_Total = Stock_Total + Cells(row, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("J" & Summary_Table_Row).Value = Stock_Letter

      ' Print the Brand Amount to the Summary Table
      Range("K" & Summary_Table_Row).Value = Stock_Total

      ' Print the Closing Stock- Opening Stock
      Range("L" & Summary_Table_Row).Value = Closing_Stock - Opening_Stock

          ' Print the i
        If Opening_Stock = 0 Then Range("M" & Summary_Table_Row).Value = 0
        Else Range("M" & Summary_Table_Row).Value = (Closing_Stock - Opening_Stock) / Opening_Stock
        End If  
      Range("M" & Summary_Table_Row).Style = "Percent"
      
      If Closing_Stock - Opening_Stock < 0 then
      Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
      End If

      If Closing_Stock - Opening_Stock > = 0 then
      Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
      End If    

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Total
      Stock_Total = 0

    End If

  Next row

End Sub
