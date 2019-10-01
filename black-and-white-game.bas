Attribute VB_Name = "Module1"
Sub playerOne(oneInput As Long):
Dim i As Long
Dim oneGg As Long
Dim twoInput As Long
Dim twoGg As Long

    ' loop for non-number entry
    ' loop for number > 99
    ' change player 1 and 2 to the names entered
    Cells(13, 3).Value = Cells(13, 3).Value + oneInput
    oneGg = 99 - Cells(13, 3).Value
    MsgBox "You have " & oneGg & " points left"
    If oneInput < 10 Then
        MsgBox "Black card"
    Else
        MsgBox "White card"
    End If
    If oneGg > 79 Then
       Range("C8:C12").Interior.Color = RGB(123, 123, 123)
    ElseIf oneGg > 59 Then
        Range("C8").Interior.Color = RGB(217, 217, 217)
    ElseIf oneGg > 39 Then
        Range("C8:C9").Interior.Color = RGB(217, 217, 217)
    ElseIf oneGg > 19 Then
        Range("C8:C10").Interior.Color = RGB(217, 217, 217)
    Else
        Range("C8:C11").Interior.Color = RGB(217, 217, 217)
    End If
    
End Sub

Sub playerTwo(twoInput As Long):
Dim i As Long
Dim oneInput As Long
Dim oneGg As Long
Dim twoGg As Long

    Cells(13, 9).Value = Cells(13, 9).Value + twoInput
    twoGg = 99 - Cells(13, 9).Value
    MsgBox "You have " & twoGg & " points left"
    If twoInput < 10 Then
        MsgBox "Black card"
    Else
        MsgBox "White card"
    End If
    If twoGg > 79 Then
       Range("I8:I12").Interior.Color = RGB(123, 123, 123)
    ElseIf twoGg > 59 Then
        Range("I8").Interior.Color = RGB(217, 217, 217)
    ElseIf twoGg > 39 Then
        Range("I8:I9").Interior.Color = RGB(217, 217, 217)
    ElseIf twoGg > 19 Then
        Range("I8:I10").Interior.Color = RGB(217, 217, 217)
    Else
        Range("I8:I11").Interior.Color = RGB(217, 217, 217)
    End If

End Sub
Sub game()

Dim one As String
Dim two As String
Dim i As Long
Dim oneInput As Long
Dim oneGg As Long
Dim twoInput As Long
Dim twoGg As Long

one = InputBox("Enter player 1 name", "Name of first player", "Player 1")
Cells(6, 3).Value = one

two = InputBox("Enter player 2 name", "Name of second player", "Player 2")
Cells(6, 9).Value = two

For i = 1 To 9
    Cells(8, 6).Value = i
    
    If i Mod 2 = 0 Then
        MsgBox two & " to enter points privately!"
        twoInput = InputBox("Enter number of points. Please only enter numbers!", "Enter points")
        Call playerTwo(twoInput)
        MsgBox one & " to enter points privately!"
        oneInput = InputBox("Enter number of points. Please only enter numbers!", "Enter points")
        Call playerOne(oneInput)
    Else
        MsgBox one & " to enter points privately!"
        oneInput = InputBox("Enter number of points. Please only enter numbers!", "Enter points")
        Call playerOne(oneInput)
        MsgBox two & " to enter points privately!"
        twoInput = InputBox("Enter number of points. Please only enter numbers!", "Enter points")
        Call playerTwo(twoInput)
    End If
    
    If oneInput > twoInput Then
        MsgBox one & " has won this round"
        Cells(11, 5).Value = Cells(11, 5).Value + 1
    Else
        MsgBox two & " has won this round"
        Cells(11, 7).Value = Cells(11, 7).Value + 1
    End If
    
    If Cells(11, 5).Value > 4 Then
        MsgBox one & " wins!"
        Exit Sub
    ElseIf Cells(11, 7).Value > 4 Then
        MsgBox two & " wins!"
        Exit Sub
    End If
Next i

End Sub

Sub reset()

Cells(11, 5).Value = 0
Cells(11, 7).Value = 0
Cells(8, 6).Value = 0
Range("C8:C12").Interior.Color = RGB(123, 123, 123)
Range("I8:I12").Interior.Color = RGB(123, 123, 123)
Cells(6, 3).Value = Null
Cells(6, 9).Value = Null
Cells(13, 3).Value = 0
Cells(13, 9).Value = 0

End Sub
