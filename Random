Sub Random()
    With ActiveSheet
        Clear
        Dim HowMany As Integer
        Dim NoOfNames As Long
        Dim RandomNumber As Integer
        Dim Names() As String 'Array to store randomly selected names
        Dim i As Byte
        Dim CellsOut As Long 'Variable to be used when entering names onto worksheet
        Dim ArI As Byte 'Variable to increment through array indexes
        Application.ScreenUpdating = False
        HowMany = Range("B1").Value - 1
        CellsOut = 3
        ReDim Names(1 To HowMany) 'Set the array size to how many names required
        NoOfNames = 21 ' Find how many names in the list
        i = 1
        Do While i <= HowMany
RandomNo:
            RandomNumber = Application.RandBetween(3, NoOfNames + 1)
            'Check to see if the name has already been picked
            For ArI = LBound(Names) To UBound(Names)
                If Names(ArI) = Cells(RandomNumber, 14).Value Then
                    GoTo RandomNo
                End If
            Next ArI
            Names(i) = Cells(RandomNumber, 14).Value ' Assign random name to the array
            i = i + 1
        Loop
        'Loop through the array and enter names onto the worksheet
        For ArI = LBound(Names) To UBound(Names)
            Cells(CellsOut, 5) = Names(ArI)
            CellsOut = CellsOut + 1
        Next ArI
        Application.ScreenUpdating = True
    End With
End Sub
