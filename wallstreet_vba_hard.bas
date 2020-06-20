Sub wallstreet_vba_hard()

For Each W In Worksheets

'set variables
    Dim lastrow As Long: lastrow = W.Cells(Rows.Count, 1).End(xlUp).Row
    Dim gp_increase As Double: gp_increase = 1
    Dim gp_decrease As Double: gp_decrease = 1
    Dim g_totalvol As Double: g_totalvol = 0

For i = 2 To lastrow
        If W.Cells(i, 11).Value > gp_increase Then
                gp_increase = W.Cells(i, 11).Value
                W.Cells(2, 16).Value = W.Cells(i, 9).Value
        End If

        If W.Cells(i, 11).Value < gp_decrease Then
                gp_decrease = W.Cells(i, 11).Value
                W.Cells(3, 16).Value = W.Cells(i, 9).Value
        End If

        If W.Cells(i, 12).Value > g_totalvol Then
                g_totalvol = W.Cells(i, 12).Value
                W.Cells(4, 16).Value = W.Cells(i, 9).Value
        End If
    Next i
        W.Cells(2, 17).Value = FormatPercent(gp_increase)
        W.Cells(3, 17).Value = FormatPercent(gp_decrease)
        W.Cells(4, 17).Value = g_totalvol
Next W

End Sub
