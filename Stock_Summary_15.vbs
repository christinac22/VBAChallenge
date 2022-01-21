Sub Stock_Summary_15()

Dim Stock_Volume As Double
Dim Summary_Row As Integer
Dim Yearly_Change As Variant
Dim Percent_Change As Variant
Dim Open_Price As Double
Dim Year_Price As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Stock Volume"

Stock_Volume = 0
Summary_Row = 2
Open_Price = Cells(2, 3).Value

For i = 2 To Range("A1").End(xlDown).Row
    
    If Cells(i, 1).Value = Cells(i + 1, 1) Then
    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
    Else
    Close_Price = Cells(i, 6).Value
    Yearly_Change = Close_Price - Open_Price
    Percent_Change = Yearly_Change / Open_Price
    
       
        If Yearly_Change > 0 Then
        Cells(Summary_Row, 10).Interior.ColorIndex = 4
        Else
        Cells(Summary_Row, 10).Interior.ColorIndex = 3
        End If
        
      
    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
    Cells(Summary_Row, 9).Value = Cells(i, 1).Value
    Cells(Summary_Row, 10).Value = Yearly_Change
    Cells(Summary_Row, 11).Value = FormatPercent(Percent_Change)
    Cells(Summary_Row, 12).Value = Stock_Volume
    
    Open_Price = Cells(i + 1, 3).Value
    Stock_Volume = 0
    Summary_Row = Summary_Row + 1
    
    End If
    
Next i

End Sub


