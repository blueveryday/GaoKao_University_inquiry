Attribute VB_Name = "Filter_Major"
Sub Filter_MajorandRange_ByKeywords()
    Dim ws As Worksheet
    Dim rngF As Range
    Dim rngH As Range
    Dim cell As Range
    Dim keywords As Variant
    Dim keyword As Variant
    Dim matchFound As Boolean
    Dim i As Integer
    Dim keywordInput As String
    Dim keywordArray() As String
    Dim defaultKeywords As String
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim userInput As String
    Dim value As Double
    Dim defaultRange As String

    ' 设置默认关键字，使用逗号分隔
    defaultKeywords = "网络,计算机,审计,海关,工科,智能,机械,航空,航天,税,数学,物理,车,电子,通信,人工,物联,软件,交通,电气,自动化,机器人,安全,测控,光电,仪器,邮政,大数据"
    
    ' 设置默认数值范围
    defaultRange = "2000,5000"
    
    ' 设置工作表
    Set ws = ThisWorkbook.Sheets("专业分数线")
    
    ' 输入关键字，使用逗号分隔
    keywordInput = InputBox("请输入关键字，用逗号分隔:", "输入关键字", defaultKeywords)
    
    ' 如果用户取消输入框，则退出子过程
    If keywordInput = "" Then Exit Sub
    
    ' 将关键字字符串拆分为数组
    keywordArray = Split(keywordInput, ",")
    
    ' 获取数值范围输入
    userInput = InputBox("请输入数值范围，用逗号分隔（例如：2000,3000）:", "输入数值范围", defaultRange)
    If userInput = "" Then Exit Sub
    On Error Resume Next
    lowerBound = CDbl(Split(userInput, ",")(0))
    upperBound = CDbl(Split(userInput, ",")(1))
    On Error GoTo 0
    If IsEmpty(lowerBound) Or IsEmpty(upperBound) Then
        MsgBox "无效的数值范围输入。", vbExclamation
        Exit Sub
    End If
    
    ' 清除所有筛选
    ws.AutoFilterMode = False
    
    ' 设置筛选范围
    Set rngF = ws.Range("F2", ws.Cells(ws.Rows.Count, "F").End(xlUp)) ' 假设 F1 是标题行
    Set rngH = ws.Range("H2", ws.Cells(ws.Rows.Count, "H").End(xlUp)) ' 假设 H1 是标题行
    
    ' 对 H 列进行升序排序
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=rngH, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 清除辅助列内容（假设 J 列和 K 列为辅助列）
    ws.Range("J:K").ClearContents
    
    ' 在标题行设置辅助列的标题
    ws.Range("J1").value = "首选专业"
    ws.Range("K1").value = "位次区间"
    
    ' 创建辅助列用于标记包含关键字的行
    For Each cell In rngF
        matchFound = False
        For i = LBound(keywordArray) To UBound(keywordArray)
            keyword = Trim(keywordArray(i))
            If InStr(cell.value, keyword) > 0 Then
                matchFound = True
                Exit For
            End If
        Next i
        If matchFound Then
            cell.Offset(0, 4).value = "包含" ' 将标记写入 J 列
        Else
            cell.Offset(0, 4).value = "不包含"
        End If
    Next cell
    
    ' 创建辅助列用于标记数值范围的行
    For Each cell In rngH
        value = cell.value
        If IsNumeric(value) Then
            If value >= lowerBound And value <= upperBound Then
                cell.Offset(0, 3).value = "在位次范围内" ' 将标记写入 K 列
            Else
                cell.Offset(0, 3).value = "不在位次范围内"
            End If
        Else
            cell.Offset(0, 3).value = "不在位次范围内"
        End If
    Next cell
    
    ' 应用筛选，先排除包含“中外合作”的行
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=6, Criteria1:="<>*中外合作*"
    
    ' 应用筛选，筛选关键字匹配的行
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=10, Criteria1:="包含"
    
    ' 应用筛选，筛选数值范围内的行
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=11, Criteria1:="在位次范围内"
    
    ' 隐藏 J 列和 K 列
    ws.Columns("J:K").Hidden = True
    
    MsgBox "专业和位次筛选完成！", vbInformation
End Sub


