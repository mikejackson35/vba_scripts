Attribute VB_Name = "Module1"

Sub StockTicker1()

Dim Ticker As String

Dim OpenPrice As Double

Dim ClosePrice As Double

Dim StockVolume As Double
    StockVolume = 0

Dim Summary_Table_Row As Long
    Summary_Table_Row = 2

Dim i As Long
Dim LastRow As Long
    LastRow = Cells(Rows.Count, 7).End(xlUp).Row
    
OpenPrice = Cells(2, 3).Value
        
For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker = Cells(i, 1).Value
        ClosePrice = Cells(i, 6).Value
        StockVolume = StockVolume + Cells(i, 7).Value
    
        
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("J" & Summary_Table_Row).Value = ClosePrice - OpenPrice
          
            If Range("J" & Summary_Table_Row).Value < 0 Then
             Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
             End If
            
            If Range("J" & Summary_Table_Row).Value >= 0 Then
             Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
             End If
             
            If OpenPrice = 0 Then
             OpenPrice = 1
             End If
             
        Range("K" & Summary_Table_Row).Value = ((ClosePrice - OpenPrice) / OpenPrice)
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        Range("L" & Summary_Table_Row).Value = StockVolume
        
        
        Summary_Table_Row = Summary_Table_Row + 1
        StockVolume = 0
        OpenPrice = Cells(i + 1, 3).Value
        
    
        
    Else
        
        StockVolume = StockVolume + Cells(i, 7).Value
        
  
    End If
    
    
Next i

    
End Sub

