Sub OnLButtonUp(Byval Item, Byval Flags, Byval x, Byval y)                 
    Dim Control, include, exclude, i
    Set Control = ScreenItems("Control")
    Set include = ScreenItems("IO-FILTER-INCLUDE")
    Set exclude = ScreenItems("IO-FILTER-EXCLUDE")

    Dim regexInclude, regexExclude
    Set regexInclude = New RegExp
    With regexInclude
        .pattern = include.InputValue
        .IgnoreCase = True 
        .Global = True
    End With

    Set regexExclude = New RegExp
    With regexExclude
        .pattern = exclude.InputValue
        .IgnoreCase = True 
        .Global = True
    End With

    For i = 0 To Control.ValueColumnCount
        Control.ValueColumnIndex = i

        if include.inputValue <> "" and exclude.inputValue = "" Then
            If regexInclude.test(Control.ValueColumnName) Or regexInclude.test(Control.ValueColumnTagName) Then
                Control.ValueColumnVisible = True
            Else
                Control.ValueColumnVisible = False
            end if
        elseif include.inputValue = "" and exclude.inputValue <> "" Then
            If regexExclude.test(Control.ValueColumnName) And regexExclude.test(Control.ValueColumnTagName) Then
                Control.ValueColumnVisible = false
            Else
                Control.ValueColumnVisible = true
            end if
        elseif include.inputValue <> "" and exclude.inputValue <> "" Then
            If (regexInclude.test(Control.ValueColumnName) Or regexInclude.test(Control.ValueColumnTagName)) And Not (regexExclude.test(Control.ValueColumnName) Or regexExclude.test(Control.ValueColumnTagName)) Then
                Control.ValueColumnVisible = true
            Else
                Control.ValueColumnVisible = false
            end if
        end if
    Next
End Sub