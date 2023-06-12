Sub CalculateStatisticsAndSort()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim rngData As Range
    Dim rngResult As Range

    ' 获取当前活动的工作簿和工作表
    Set wb = ThisWorkbook
    Set wsSource = wb.ActiveSheet
    
    ' 输入要处理的数据范围
    On Error Resume Next
    Set rngData = Application.InputBox("请输入要处理的数据范围:", Type:=8)
    On Error GoTo 0
    
    ' 如果用户取消了输入，则退出子程序
    If rngData Is Nothing Then Exit Sub
    
    ' 创建处理结果表
    On Error Resume Next
    Set wsResult = wb.Sheets("处理结果")
    On Error GoTo 0
    
    If wsResult Is Nothing Then
        Set wsResult = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsResult.Name = "处理结果"
    End If
    
    ' 设置处理结果表的单元格
    Set rngResult = wsResult.Range("A1:A7")
    
    ' 计算统计数据
    rngResult.Cells(1).Value = WorksheetFunction.Max(rngData)
    rngResult.Cells(3).Value = WorksheetFunction.Min(rngData)
    rngResult.Cells(5).Value = WorksheetFunction.Median(rngData)
    rngResult.Cells(7).Value = WorksheetFunction.Mode(rngData)
    
    ' 将数据进行升序排序
    rngData.Copy
    wsResult.Range("B1").PasteSpecial Paste:=xlPasteValues
    wsResult.Range("B1:B" & rngData.Rows.Count).Sort Key1:=wsResult.Range("B1"), Order1:=xlAscending, Header:=xlNo
    
    ' 清除剪贴板中的内容
    Application.CutCopyMode = False
    
    ' 输入文字到指定单元格
    wsResult.Range("A2").Value = "最大值"
    wsResult.Range("A4").Value = "最小值"
    wsResult.Range("A6").Value = "平均数"
    wsResult.Range("A8").Value = "众数"
    wsResult.Range("C1").Value = "排序结果"
    
    ' 提示完成处理
    MsgBox "数据处理完成！"
End Sub