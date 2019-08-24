VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub walstreet_data():
   Dim totalvol As Double
   Dim j As Integer
   Dim yearlychange As Double
   Dim percentchange As Double
   Dim openingprice As Double
   Dim closingprice As Double
   Dim openingprice_row As Long
       totalvol = 0
       j = 2
       openingprice_row = 2
       Range("I1").Value = "Ticker"
       Range("J1").Value = "Total Volume"
       Range("K1").Value = "Yearly Change"
       Range("L1").Value = "Percent Change"
       lastrow = Cells(Rows.Count, "A").End(xlUp).Row
       For i = 2 To lastrow
          If Range("A" & i + 1).Value = Range("A" & i).Value Then
               totalvol = totalvol + Range("G" & i).Value
           Else
               openingprice = Range("C" & openingprice_row)
               closingprice = Range("F" & i)
               yearlychange = closingprice - openingprice
               If openingprice <> 0 Then
                   percentchange = yearlychange / openingprice
               End If
               Range("I" & j).Value = Range("A" & i).Value
               Range("J" & j).Value = totalvol + Range("G" & i).Value
               Range("K" & j).Value = yearlychange
               Range("K" & j).NumberFormat = "0.00000000"
               Range("L" & j).Value = percentchange
               Range("L" & j).NumberFormat = "0.00%"
               If Range("K" & j).Value > 0 Then
                   Range("K" & j).Interior.ColorIndex = 4
               Else
                   Range("K" & j).Interior.ColorIndex = 3
               End If
               j = j + 1
               totalvol = 0
               openingprice_row = i + 1
           End If
       Next i
End Sub




