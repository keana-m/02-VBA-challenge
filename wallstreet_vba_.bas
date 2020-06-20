Sub wallstreet_vba()

For Each W In Worksheets

'set variables
    Dim lastrow As Long: lastrow = W.Cells(Rows.Count, 1).End(xlUp).Row
    Dim ticker As String
    Dim opendate As Double
    Dim closedate As Double
    Dim yearlychange As Double
    Dim percentchange As Variant
    Dim volumetotal As Double: volumetotal = 0
    Dim stockrow As Integer: stockrow = 2
    Dim i_start As Double: i_start = 2

'retrieve tables
    W.Cells(1, 9).Value = "Ticker"
    W.Cells(1, 10).Value = "Yearly Change"
    W.Cells(1, 11).Value = "Percent Change"
    W.Cells(1, 12).Value = "Total Stock Volume"
    W.Cells(2, 15).Value = "Greatest % Increase"
    W.Cells(3, 15).Value = "Greatest % Decrease"
    W.Cells(4, 15).Value = "Greatest Total Stock Volume"
    W.Cells(1, 17).Value = "Value"
    W.Cells(1, 16).Value = "Ticker"

'start loop
    For i = 2 To lastrow
    ticker = W.Cells(i, 1).Value
    volumetotal = volumetotal + W.Cells(i, 7).Value
        If W.Cells(i + 1, 1).Value <> ticker Then
                opendate = W.Cells(i_start, 3).Value
                closedate = W.Cells(i, 6).Value
                yearlychange = (closedate - opendate)
                percentchange = 0
            
            If opendate <> 0 Then
                percentchange = (closedate / opendate) - 1
            End If
            
'--------------------------------------------------------------------------------
    volumetotal = volumetotal + W.Cells(i, 7).Value
        If W.Cells(i + 1, 1).Value <> W.Cells(i, 1).Value Then
            W.Cells(stockrow, 12).Value = volumetotal
        End If
    W.Cells(stockrow, 9).Value = ticker
    W.Cells(stockrow, 10).Value = yearlychange

            If yearlychange < 0 Then
                W.Cells(stockrow, 10).Interior.Color = vbRed
                Else: W.Cells(stockrow, 10).Interior.Color = vbGreen
            End If
'--------------------------------------------------------------------------------
            
    W.Cells(stockrow, 11).Value = FormatPercent(percentchange)
    W.Cells(stockrow, 12).Value = volumetotal
    stockrow = stockrow + 1
    i_start = i + 1
    ticker = W.Cells(i + 1, 1).Value
    volumetotal = 0
      
      End If
    Next i
Next W
End Sub


