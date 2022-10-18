Private Sub CommandButton1_Click()

    ' -- declaration
    Dim rng As Range, crng As Range
    Dim n As Long
    Dim t As Long
    Dim f As Double
    Dim pth As String
    Dim nm As String
    Dim c As Long
    Dim rowPerBatch As Long
    Dim nbOfBatch As Long
    Dim wb As Workbook, cpy As Workbook
    Dim rstart As Long, rend As Long
    
    ' -- initialization
    Set rng = Range(UserForm1.RefEdit1.Value)
    n = CLng(UserForm1.TextBox1.Value)
    Set wb = ActiveWorkbook
    pth = wb.Path & "\"
    nm = wb.Name
    
    ' -- determine how many rows per batch
    If UserForm1.OptionButton1.Value = True Then
        ' the user splits by number of lines
        rowPerBatch = n
        t = rng.Rows.Count
        nbOfBatch = Application.WorksheetFunction.RoundDown(t / rowPerBatch, 0)
        If t Mod rowPerBatch > 0 Then nbOfBatch = nbOfBatch + 1
    Else
        ' the user splits by number of batches
        nbOfBatch = n
        t = rng.Rows.Count
        rowPerBatch = Application.WorksheetFunction.RoundUp(t / n, 0)
    End If
    
    ' -- split
    For c = 1 To nbOfBatch
        wb.SaveCopyAs pth & Right("00" & CStr(c), 3) & "_" & nm
        Set cpy = Workbooks.Open(pth & Right("00" & CStr(c), 3) & "_" & nm)
        Set crng = cpy.Worksheets(rng.Worksheet.Index).Range(rng.Address)
        rstart = getStart(crng.Cells(1).Row, rowPerBatch, c)
        rend = rstart + rowPerBatch - 1
        ' If rend > crng.Cells(crng.Cells.Count).Row Then rend = crng.Cells(crng.Cells.Count).Row
        ' delete all rows after (as long as they belong to the original selected range)
        If c < nbOfBatch Then
            cpy.Worksheets(rng.Worksheet.Index).Range("A" & CStr(rend + 1) & ":A" & CStr(crng.Cells(crng.Cells.Count).Row)).EntireRow.Delete
        End If
        ' delete all rows before (as long as they belong to the original selected range)
        If c > 1 Then
            cpy.Worksheets(rng.Worksheet.Index).Range("A" & CStr(crng.Cells(1).Row) & ":A" & CStr(rstart - 1)).EntireRow.Delete
        End If
        cpy.Save
        cpy.Close
    Next c
    
End Sub

Function getStart(rowStart As Long, batchSize As Long, batchNb As Long) As Long
   Dim r As Long
   r = rowStart + ((batchNb - 1) * batchSize)
   getStart = r
End Function
