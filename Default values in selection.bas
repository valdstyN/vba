' Change all blank values in the cell selection to a default value
Sub defaultValue()
    Dim rng As Range
    Dim newval As String
    newval = InputBox("Replace blank values with?")
    For Each rng In Application.Selection
        If rng.Value = "" Then rng.Value = newval
    Next rng
End Sub
