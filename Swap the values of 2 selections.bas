Sub swapCol()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.EnableEvents = False

    Dim r1 As Range, r2 As Range, r3 As Variant, n As Range, i As Long

    Set r1 = Application.Selection.Areas.Item(1)
    Set r2 = Application.Selection.Areas.Item(2)
    Set r1p = Application.Selection.Areas.Item(1)
    r3 = r1
    r2.Copy r1

    For i = 1 To UBound(r3)
     r2(i, 1).Value = r3(i, 1)
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
End Sub
