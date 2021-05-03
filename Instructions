Sub Instructions()
    With ActiveSheet
        If Range("B10").Value = "" Then
            MsgBox "Fill out the number of restaurants, the mode of transportation, and time limit in hours in cells B1:B3. Then in Column E, fill out the restaurants that you want to go to in the yellow cells. Northeastern (NEU) must go first. If you are feeling lazy, click on 'Random Restaurants' to pick for you. Then click generate routes."
        Else
            MsgBox "Your table of paths have been generated in the gray cells. To find the optimal path, go to Solver in the Data section. The Solver has already been set up for you. Check the blue cells and make sure the cell range for 'changing variable cells' and the 'allDifferent constraint' is the same as the ones in the blue cells in column H."
            MsgBox "The optimal path will be in row 18, the restaurant row, from left to right. Check cells B6:B7 to see if your optimal route fits within your time limits. If you are done or wish to find the optimal path for a new set of restaurants, click the clear button and start again."
        End If
    End With
End Sub
