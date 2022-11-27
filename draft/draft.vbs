Sub arr()

    ' enable Microsoft Scripting Runtime
    Dim emp As Dictionary
    Set emp = New Dictionary
    
    ' --- FIRST PART OF THE EXECUTION WHERE WE USED TO DETERMINE THAT AN EMPLOYEE NEEDS TO BE FLAGGED "X"
    ' --- NOW INSTEAD OF A FLAG "X", WE HAVE A HASHMAP FOR EMPLOYEES.
    ' --- EACH RECORD CONTAINS THE LIST OF SCENARIOS WITH EFFECTIVE DATE
    ' --- WE DON'T CAPTURE THE POPULATION THAT WAS APPLICABLE THEN
    
    'for each e in employee
    '  create hashmap
        emp("101") = Array()
    '   for each f in filter
    '   ...check all criteria via EVAL
    '   if match = true then
    '       add scenario and date to hashmap of employee
            emp("101") = addItem(emp("101"), "01/01/1999;SC2")
            ' if we matched another set of criteria, we'd apply
            emp("101") = addItem(emp("101"), "01/01/2001;SC3")
    '   end if
    '   next f filter
    'next e
    
    ' ubound(emp("xxx")) = last index

    ' --- SECOND PART OF THE EXECUTION WHERE WE APPLY SCENARIOS
    'for each e in employee
        ' eid = e.value
        ' for each DATE (header)
        ' -- HERE, WE ARE PROCESSING 1 EMPLOYEE AND NEED TO KNOW WHICH SCENARIO TO APPLY AT THE GIVEN DATE (say DAT) (eg 01 09 2022, 01 09 2023...)
        ' -- WE HAVE TO CHECK IN THE HASHMAP : either we find the exact date, a segment, or it may not be in - in that case default)
        dat = CDate(thisDate)
        curDt = CDate("01/09/2022") ' default date -  should be the first date
        curSc = "/"                 ' default scenario
        For Each dt In emp("101")
            thisDt = Split(emp("101"), ";")(0)
            thisSc = Split(emp("101"), ";")(1)
            If CDate(thisDt) >= dat Then
                curDt = thisDt
                curSc = thisSc
                GoTo applySc    ' we stop as soon as we find a scenario to apply with a higher/equal date
            End If
            
applySc:
        Next dt
    'next e

End Sub

Function addItem(arr As Variant, newStr As String) As Variant
    Dim a As Variant
    a = arr
    ReDim Preserve a(UBound(arr) + 1)
    a(UBound(arr) + 1) = newStr
    addItem = a
End Function
