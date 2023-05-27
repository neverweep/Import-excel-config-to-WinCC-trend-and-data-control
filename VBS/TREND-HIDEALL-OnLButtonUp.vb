Sub OnLButtonUp(Byval Item, Byval Flags, Byval x, Byval y)        
	Dim Control,i
	Set Control = ScreenItems("Control")

	For i = 0 To Control.TrendCount
		Control.TrendIndex = i
		Control.TrendVisible = False
	Next
End Sub