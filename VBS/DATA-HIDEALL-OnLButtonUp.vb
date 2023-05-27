Sub OnLButtonUp(Byval Item, Byval Flags, Byval x, Byval y)         
Dim Control,i
Set Control = ScreenItems("Control")

For i = 0 To Control.ValueColumnCount
    Control.ValueColumnIndex = i
    Control.ValueColumnVisible = False
Next
End Sub