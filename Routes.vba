Sub Routes()
    With ActiveSheet
        If Range("E2").Value <> "NEU" Then
            MsgBox "First Restaurant must be NEU"
            End
        ElseIf Range("M2").Value <> Range("B1").Value Then
            Dim msg_string
            
            If Range("M2").Value = 1 Then
                msg_string = "You have filled out " & Range("M2").Value & " restaurant. You need to fill out " & Range("B1").Value & " restaurants."
                MsgBox msg_string
                End
            End If
            
            msg_string = "You have filled out " & Range("M2").Value & " restaurants. You need to fill out " & Range("B1").Value & " restaurants."
            MsgBox msg_string
            End
        End If
        Dim i
        Dim j
        i = 1
        While i <= Range("M2").Value
            j = 1
            While j <= Range("M2").Value
                
                If i = j Then
                    j = j + 1
                End If
                
                If Cells(j + 1, 5).Value = Cells(i + 1, 5).Value Then
                    MsgBox "You have a duplicate restaurant"
                    End
                End If
                j = j + 1
            Wend
            i = i + 1
        Wend
                    
        Dim Counter
        Counter = 1
        While (Counter <= Range("B1").Value)
            'Columns
            .Cells(Counter + 10, 1) = .Cells(Counter + 1, 5)
            .Cells(Counter + 10, 2) = Counter
            'Rows
            .Cells(9, Counter + 2) = .Cells(Counter + 1, 5)
            .Cells(10, Counter + 2) = Counter
            
            'Solution
            .Cells(19, Counter + 2) = Counter
                    
            Counter = Counter + 1
        Wend
        
        .Cells(19, Range("B1").Value + 3) = 1
        .Range("B10").Value = "Row/Column Number"
        
        If Range("B1").Value = 3 Then
            .Range("C20").Value = "=INDEX(C11:H16, C19, D19)"
            .Range("D20").Value = "=INDEX(C11:H16, D19, E19)"
            .Range("E20").Value = "=INDEX(C11:H16, E19, F19)"
            .Range("F20").Value = "N/A"
            .Range("H6:H7").Value = "$C$19:$E$19"
        ElseIf Range("B1").Value = 4 Then
            .Range("C20").Value = "=INDEX(C11:H16, C19, D19)"
            .Range("D20").Value = "=INDEX(C11:H16, D19, E19)"
            .Range("E20").Value = "=INDEX(C11:H16, E19, F19)"
            .Range("F20").Value = "=INDEX(C11:H16, F19, G19)"
            .Range("G20").Value = "N/A"
            .Range("H6:H7").Value = "$C$19:$F$19"
        ElseIf Range("B1").Value = 5 Then
            .Range("C20").Value = "=INDEX(C11:H16, C19, D19)"
            .Range("D20").Value = "=INDEX(C11:H16, D19, E19)"
            .Range("E20").Value = "=INDEX(C11:H16, E19, F19)"
            .Range("F20").Value = "=INDEX(C11:H16, F19, G19)"
            .Range("G20").Value = "=INDEX(C11:H16, G19, H19)"
            .Range("H20").Value = "N/A"
            .Range("H6:H7").Value = "$C$19:$G$19"
        ElseIf Range("B1").Value = 6 Then
            .Range("C20").Value = "=INDEX(C11:H16, C19, D19)"
            .Range("D20").Value = "=INDEX(C11:H16, D19, E19)"
            .Range("E20").Value = "=INDEX(C11:H16, E19, F19)"
            .Range("F20").Value = "=INDEX(C11:H16, F19, G19)"
            .Range("G20").Value = "=INDEX(C11:H16, G19, H19)"
            .Range("H20").Value = "=INDEX(C11:H16, H19, I19)"
            .Range("I20").Value = "N/A"
            .Range("H6:H7").Value = "$C$19:$H$19"
        End If
        MsgBox "Use Excel Solver to find the optimal path. Click Ok to read the instructions in blue."
    End With
End Sub
