Sub Clear()
    With ActiveSheet
        .Range("C9:H10").ClearContents
        .Range("A11:B16").ClearContents
        .Range("E3:E7").ClearContents
        .Range("B10").ClearContents
        .Range("E2").Value = "NEU"
        .Range("C19:J20").ClearContents
        .Range("H6:H7").ClearContents
    End With
End Sub
