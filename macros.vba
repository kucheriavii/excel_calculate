Sub calculate()
For i = 2 To 40
    If Cells(i, 4).Value > 0 And Cells(i, 4).Value <> "" Then
        For j = 2 To 40
            If (Cells(i, 4).Value > Cells(j, 12).Value) Then
                Cells(i, 8) = Cells(j + 1, 13)
            End If
            If (Cells(i, 4).Value = Cells(j, 12).Value) Then
                Cells(i, 8) = Cells(j, 13)
            End If
        Next
    End If
    
Next
End Sub
