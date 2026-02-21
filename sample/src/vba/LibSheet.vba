Function NewWorksheet(name As String, Optional wb As Workbook = Nothing) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    On Error GoTo ErrorHandler
    Set ws = wb.Worksheets(name)
    On Error GoTo 0

    ws.Cells.Delete shift:=xlUp

    Set NewWorksheet = ws
    Exit Function

ErrorHandler:
    Set ws = wb.Worksheets.Add
    ws.Name = name
    Resume Next
End Function
