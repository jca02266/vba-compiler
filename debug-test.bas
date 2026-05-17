Class Container
    Private val
    Property Get Value()
        Value = val
    End Property
    Property Let Value(v)
        val = v
    End Property
End Class

Sub ModifyValue(c As Container)
    c.Value = c.Value + 10
End Sub

Function Test()
    Dim x As New Container
    x.Value = 5
    ModifyValue x
    Test = x.Value
End Function
