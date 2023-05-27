Sub OnLButtonUp(Byval Item, Byval Flags, Byval x, Byval y)                 
	Dim Control,i
	Set Control = ScreenItems("Control")

	Dim regex
	Set regex = New RegExp
	With regex
		.pattern = "(温度|Temp.|Temperature)" '使用正则表达式匹配需要显示的趋势名(TrendName)
		.IgnoreCase = True 
		.Global = True
	End With

	For i = 0 To Control.TrendCount
		Control.TrendIndex = i
		If regex.test(Control.TrendName) Then 
			Control.TrendVisible = True
		Else
			Control.TrendVisible = False
		End If
	Next
End Sub