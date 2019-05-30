Attribute VB_Name = "Module1"
Sub Stock()

 ' Set an initial variable for holding the Ticker and worksheet
 
 Dim ws As Worksheet
 For Each ws In Worksheets
 ws.Activate
 Dim Ticker As String

 ' Set an initial variable for holding the total stock volume
 Dim Tota1_stock_volume As Double
 Total_Stock_Volume = 0
 Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2
  ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Totalstockvolume"


   ' Loop through all the rows

       lastrow = Cells(Rows.Count, 1).End(xlUp).Row
       For i = 2 To lastrow

        ' Check if we are still within the same credit card brand, if it is not...
          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

              ' Set the Ticker value
                Ticker = Cells(i, 1).Value

              ' Add to the  Total_stock_volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

              ' Print the Ticker value in the Summary Table
                 ws.Range("I" & Summary_Table_Row).Value = Ticker

              ' Print the Total stock volume to the Summary Table
                 ws.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

              ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

             ' Reset the Total stock volume
                Total_Stock_Volume = 0

          ' If the cell immediately following a row is the same brand...
         Else

               ' Add to the Brand Total
               Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

         End If

 Next i

Next ws

End Sub
