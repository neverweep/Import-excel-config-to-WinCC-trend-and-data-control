Sub OnOpen()                                      
    Dim Control
    Dim ADOConn, ADORecordset
    Dim xlsPath, xlsSheet
    Dim i

    Set ADOConn = CreateObject("ADODB.Connection")
    Set ADORecordset = CreateObject("ADODB.Recordset")
    Set Control = ScreenItems("Control")

    '循环清空当前表格控件内容
    'For i = 0 To Control.ValueColumnCount
    '    Control.ValueColumnIndex = i
    '    Control.ValueColumnRemove = Control.ValueColumnName
    'Next
    'i = 0


    xlsPath = HMIRuntime.ActiveProject.Path & "\GraCS\Archive.xls" '获取excel文件绝对路径
    xlsSheet = "Sheet1" '打开excel的sheet1页

    Control.Online = True '每次打开都设置为在线，保证显示最新数据

    HMIRuntime.Trace "Data archive configration ^^^^^" & vbCrLf '输出调试日志

    If Control.ValueColumnCount = 0 Then '如果当前表格控件条目数为0

        HMIRuntime.Trace "Configration not exsits" & vbCrLf

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
            If InStr(1,ADORecordset.fields(7).value, "N", 1) <= 0 Then
            Control.ValueColumnadd = ADORecordset.fields(0).value
                If i <= 2 Then
                Control.ValueColumnVisible = True
                    i=i+1
                Else
                Control.ValueColumnVisible = False
                End If
            Control.ValueColumnTagName = ADORecordset.fields(8).value & ADORecordset.fields(1).value '将数据应用到Tag
            Control.ValueColumnPrecisions  = ADORecordset.fields(10).value '将数据应用到小数位数
            Control.ValueColumnAutoPrecisions  = False '取消自动小数
            Control.ValueColumnTimeColumn = ADORecordset.fields(9).value '将数据应用到时间列
            Control.ValueColumnForeColor = RGB(256 * Rnd, 256 * Rnd, 256 * Rnd) '给表格应用不同颜色
            Control.ValueColumnLength = 24 '设置列宽
            Control.ValueColumnAlign = 2 '设置列对齐方式
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
    HMIRuntime.Trace "Data archive configration ^^^^^" & vbCrLf & vbCrLf
End Sub