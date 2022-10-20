Sub StripAccent()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.EnableEvents = False

    Dim n As Range
    Dim A As String * 1
    Dim B As String * 1
    Dim i As Integer

    Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
    Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
  
    For Each n In Application.Selection
        For i = 1 To Len(AccChars)
            A = Mid(AccChars, i, 1)
            B = Mid(RegChars, i, 1)
            n.Value = Replace(n.Value, A, B)
        Next
    Next n

    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
End Sub
