Sub OnLButtonUp(Byval Item, Byval Flags, Byval x, Byval y)
    Dim Control,
    Dim ADOConn, ADORecordset
    Dim xlsPath, xlsSheet
    Dim i

    Set ADOConn = CreateObject("ADODB.Connection")
    Set ADORecordset = CreateObject("ADODB.Recordset")
    Set Control = ScreenItems("Control")

    '循环清空当前趋势控件内容
    For i = 0 To Control.TrendCount
        Control.TrendIndex = i
        Control.TrendRemove = Control.TrendName
    Next
    i = 0

    xlsPath = HMIRuntime.ActiveProject.Path & "\GraCS\Archive.xls" '获取excel文件绝对路径
    xlsSheet = "Sheet1" '打开excel的sheet1页

    Control.Online = True '每次打开都设置为在线，保证显示最新数据
    Control.ValueAxisAutoRange = True 
    Control.ValueAxisAutoRange = False '重置自动范围，保证每次打开显示范围正确

    HMIRuntime.Trace "Trend archive configration ^^^^^" & vbCrLf '输出调试日志

    If Control.TrendCount = 0 Then '如果当前趋势控件条目数为0

        HMIRuntime.Trace "Configration not exists" & vbCrLf

        '没有发现excel文件，则报异常错误
        On Error Resume Next
        ADOConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xlsPath & ";Extended Properties=Excel 8.0;"
        ADORecordset.Open "Select * from [" & xlsSheet & "$]", ADOConn, 3, 3
        If err.number <> 0 Then
            HMIRuntime.Trace "Please check: " & xlsPath & vbCrLf
            Exit Sub
        End If
        
        '将数据移动到excel文件第三行
        ADORecordset.MoveFirst
        ADORecordset.MoveNext

        '如果没有到excel行尾，则重复循环
        Do While Not ADORecordset.BOF And Not ADORecordset.EOF
            If InStr(1,ADORecordset.fields(3).value, "N", 1) <= 0 Then
                Control.Trendadd = ADORecordset.fields(0).value
                If i <= 2 Then
                    Control.TrendVisible = True
                    i=i+1
                Else
                    Control.TrendVisible = False
                End If
                Control.TrendTagName = ADORecordset.fields(3).value & ADORecordset.fields(1).value '将数据应用到趋势名
                Control.TrendTrendWindow = ADORecordset.fields(4).value '将数据应用到趋势窗口名
                Control.TrendTimeAxis = ADORecordset.fields(5).value '将数据应用到时间轴名
                Control.TrendValueAxis = ADORecordset.fields(6).value '将数据应用到数值轴名
                Control.TrendColor = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd) '给趋势应用不同颜色
            End If
            ADORecordset.MoveNext '打开下一条数据记录
        Loop

        '关闭数据连接，关闭文件，释放对象
        ADORecordset.Close
        ADOConn.Close
        Set ADORecordset = Nothing
        Set ADOConn = Nothing
        
        HMIRuntime.Trace "Successfully loaded archive configration" & vbCrLf
    Else
        HMIRuntime.Trace "Configration already exists" & vbCrLf
    End If
    HMIRuntime.Trace "Trend archive configration ^^^^^" & vbCrLf & vbCrLf
End Sub