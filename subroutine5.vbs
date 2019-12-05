

Sub subroutine5()
For i = 1 To 4
    If Sheets("Sheet1").Cells([i], [2]).Value = "Michael" Then  'Cells([row], [col]).Value or .Text
        'Perform your action here
        '-4105 for default color, RGBY(std and light): 3 & 24, 5 & 17, 4 & 50, 6 & 36
        Sheets("Sheet2").Cells([i], [1]).Value = Sheets("Sheet1").Cells([i], [1]).Value
    End If
Next i
End Sub


Sub subroutine5()
For i = 1 To 4
    If Left(Split(Sheets("Sheet1").Cells([i], [2]).Value, " ")(0), 1) = "1" Or Left(Split(Sheets("Sheet1").Cells([i], [2]).Value, " ")(0), 1) = "4" Then
        'Perform your action here
        '-4105 for default color, RGBY(std and light): 3 & 24, 5 & 17, 4 & 50, 6 & 36
        Sheets("Sheet2").Cells([i], [1]).Value = Sheets("Sheet1").Cells([i], [1]).Value
    End If
Next i
End Sub


Sub subroutine5()
For i = 1 To 4
    If Trim(Left(Split(Sheets("Sheet1").Cells([i], [2]).Value, " ")(0), 1)) = "1" Or Trim(Left(Split(Sheets("Sheet1").Cells([i], [2]).Value, " ")(0), 1)) = "4" Then
        'Perform your action here
        '-4105 for default color, RGBY(std and light): 3 & 24, 5 & 17, 4 & 50, 6 & 36
        Sheets("Sheet2").Cells([i], [1]).Value = Sheets("Sheet1").Cells([i], [1]).Value
    End If
Next i
End Sub


