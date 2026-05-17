Function NewWorkbook(Optional SheetCount As Integer = 1) As Workbook
    Dim sv As Integer

    sv = Application.SheetsInNewWorkbook
    On Error GoTo Cleanup
    Application.SheetsInNewWorkbook = SheetCount
    Set NewWorkbook = Workbooks.Add

Cleanup:
    Application.SheetsInNewWorkbook = sv
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function
