Sub OnLButtonUp(Byval Item, Byval Flags, Byval x, Byval y)                  
    Dim Control,i
    Set Control = ScreenItems("Control")

    Dim regex
    Set regex = New RegExp
    With regex
        .pattern = "(温度|Temp.|Temperature)" '使用正则表达式匹配需要显示的表格名(ValueColumnName)
        .IgnoreCase = True 
        .Global = True
    End With

    For i = 0 To Control.ValueColumnCount
        Control.ValueColumnIndex = i
        If regex.test(Control.ValueColumnName) Then 
            Control.ValueColumnVisible = True
        Else
            Control.ValueColumnVisible = False
        End If
    Next
End Sub