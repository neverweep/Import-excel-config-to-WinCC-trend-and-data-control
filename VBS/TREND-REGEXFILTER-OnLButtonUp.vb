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

    For i = 0 To Control.TrendCount
        Control.TrendIndex = i

        if include.inputValue <> "" and exclude.inputValue = "" Then
            If regexInclude.test(Control.TrendName) Or regexInclude.test(Control.TrendTagName) Then
                Control.TrendVisible = True
            Else
                Control.TrendVisible = False
            end if
        elseif include.inputValue = "" and exclude.inputValue <> "" Then
            If regexExclude.test(Control.TrendName) And regexExclude.test(Control.TrendTagName) Then
                Control.TrendVisible = false
            Else
                Control.TrendVisible = true
            end if
        elseif include.inputValue <> "" and exclude.inputValue <> "" Then
            If (regexInclude.test(Control.TrendName) Or regexInclude.test(Control.TrendTagName)) And Not (regexExclude.test(Control.TrendName) Or regexExclude.test(Control.TrendTagName)) Then
                Control.TrendVisible = true
            Else
                Control.TrendVisible = false
            end if
        end if
    Next
End Sub