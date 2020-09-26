Sub creditCard()

Dim card As String
Dim amount As Double
Dim summary As Integer

For i = 2 To 101
    
    card = Cells(i, 1).Value
    amount = Cells(i, 3).Value

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    Cells(2, 7).Value = card
    Cells(2, 8).Value = amount

    
    Else
    
    amount = amount + Cells(i, 3).Value
   
    
    End If
    

Next i


End Sub

