' example
' sql(qselect:="[First Name],[Last Name]", qfrom:="Sheet1", qwhere="[EmpID]=101")
Function sql(qselect As String, qfrom As String, qwhere As String) As Variant
    Dim r As Variant
    Set r = CreateObject("System.Collections.ArrayList")
    Dim sh As Worksheet, rng As Range, rng2 As Range, a As String, b As Boolean, c As Long, f As String
    ' set FROM worksheet
    Set sh = ThisWorkbook.Sheets(qfrom)
    ' set WHERE column
    ' find column
    a = Split(qwhere, "]")(0)
    a = Mid(a, 2, Len(a))
    For Each rng In sh.Range("A1:ZZ1")
        If rng.Value = a Then
            a = Replace(rng.Address, "$1", "")
            a = a & ":" & a
            Exit For
        End If
    Next rng
    ' get SELECT fields
    b = False
    For Each rng In sh.Range(a)
        If b = True Then
            If Evaluate(rng.Value & Mid(qwhere, InStr(qwhere, "]") + 1)) Then
                ' found, hence retrieve all values from qselect parameter
                For c = 0 To UBound(Split(qselect, ","))
                    f = Split(qselect, ",")(c)
                    f = Mid(f, 2, Len(f) - 2)
                    For Each rng2 In sh.Range("A1:ZZ1")
                        If rng2.Value = f Then
                            r.Add rng.Offset(0, rng2.Column - rng.Column).Value
                        End If
                    Next rng2
                Next c
                Exit For ' BECAUSE OF THIS, WE ARE STOPPING AT THE FIRST ENTRY FOUND
            End If
        End If
        b = True
    Next rng
    Set sql = r
End Function


Sub test()
    Dim r As Variant
    Set r = sql(qselect:="[First Name],[Last Name],[Seniority Date]", qfrom:="emp", qwhere:="[Grade]=2")
    Dim i As Long
    For i = 0 To r.Count - 1
        Debug.Print r(i)
    Next i
End Sub
