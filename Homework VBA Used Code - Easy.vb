Sub RunLoop():

 Dim iIndex As Integer
    Dim ws As Excel.Worksheet

    For iIndex = 1 To ActiveWorkbook.Worksheets.Count
        Set ws = Worksheets(iIndex)
        ws.Activate

        


  ' Set an initial variable for holding the Ticker name
  Dim Ticker As String
  Dim Total_Stock_volume As Double
  Total_Stock_volume = 0
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Set sht = ActiveSheet
  lastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
  
        For i = 2 To lastRow
            Range("I1") = "Ticker"
            Range("J1") = "Total Stock Volume"
    ' Check if we are still within the same ticker, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

             ' Set the Ticker name
            Ticker = Cells(i, 1).Value

      ' Add to the  Total
            Total_Stock_volume = Total_Stock_volume + Cells(i, 7).Value

      ' Print the ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the ticker Amount to the Summary Table
            Range("J" & Summary_Table_Row).Value = Total_Stock_volume

      ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the  Total
            Total_Stock_volume = 0

    ' If the cell immediately following a row is the same ticker...
            Else

      ' Add to the  Total
            Total_Stock_volume = Total_Stock_volume + Cells(i, 7).Value

            End If
        Next i



    Next iIndex
End Sub





