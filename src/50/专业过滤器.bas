Attribute VB_Name = "专业过滤器"
Sub A_Filter_MajorandRange_ByKeywords()
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
    Dim excludeKeywordInput As String
    Dim excludeKeywordArray() As String
    Dim lowerBound As Double
    Dim upperBound As Double
    Dim userInput As String
    Dim value As Variant
    Dim yearInput As String
    Dim categoryInput As String
    Dim defaultCategory As String
    Dim batchInput As String
    Dim defaultBatch As String
    Dim excludeKeywords As String
    
    defaultKeywords = "网络,计算机,审计,海关,工科,智能,机械,航空,航天,税,数学,物理,电子,通信,人工,物联,软件,交通,电气,自动化,机器人,安全,测控,光电,仪器,邮政,大数据,船舶"
    defaultRange = "2900,6000"
    excludeKeywords = "中外,合作,民族"

    Set ws = ThisWorkbook.Sheets("专业分数线")

    yearInput = InputBox("请输入要筛选的招生年份（例如：2023或2024）:", "输入招生年份", "2024")
    If yearInput = "" Then Exit Sub

    ws.AutoFilterMode = False
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:=yearInput

    defaultCategory = "物理类"
    categoryInput = InputBox("请输入要筛选的类别（物理类 或 历史类）:", "输入类别", defaultCategory)
    If categoryInput = "" Then Exit Sub
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=4, Criteria1:=categoryInput

    defaultBatch = "本科批"
    batchInput = InputBox("请输入要筛选的批次（本科批、本科提前批、本科提前批B段）:", "输入批次", defaultBatch)
    If batchInput = "" Then Exit Sub
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=5, Criteria1:=batchInput

    Set rngF = ws.Range("F2", ws.Cells(ws.Rows.Count, "F").End(xlUp))
    Set rngH = ws.Range("H2", ws.Cells(ws.Rows.Count, "H").End(xlUp))

    keywordInput = InputBox("请输入关键字，用逗号分隔:", "输入关键字", defaultKeywords)
    If keywordInput = "" Then Exit Sub
    keywordArray = Split(keywordInput, ",")

    excludeKeywordInput = InputBox("请输入要排除的专业名称关键字，用逗号分隔:", "输入排除关键字", excludeKeywords)
    If excludeKeywordInput = "" Then excludeKeywordArray = Split(excludeKeywords, ",") Else excludeKeywordArray = Split(excludeKeywordInput, ",")

    If excludeKeywordInput = "" Then Exit Sub

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

    ws.Range("J:K").ClearContents
    ws.Range("J1").value = "首选专业"
    ws.Range("K1").value = "位次区间"

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
            cell.Offset(0, 4).value = "包含"
        Else
            cell.Offset(0, 4).value = "不包含"
        End If
    Next cell

    For Each cell In rngF
        For i = LBound(excludeKeywordArray) To UBound(excludeKeywordArray)
            Dim excludeKeyword As String
            excludeKeyword = Trim(excludeKeywordArray(i))
            If InStr(cell.value, excludeKeyword) > 0 Then
                cell.Offset(0, 4).value = "排除"
                Exit For
            End If
        Next i
    Next cell

    For Each cell In rngH
        If IsNumeric(cell.value) Then
            value = cell.value
            If value >= lowerBound And value <= upperBound Then
                cell.Offset(0, 3).value = "在位次范围内"
            Else
                cell.Offset(0, 3).value = "不在位次范围内"
            End If
        Else
            cell.Offset(0, 3).value = "不在位次范围内"
        End If
    Next cell

    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=6, Criteria1:="<>*中外合作*"
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=10, Criteria1:="包含"
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=11, Criteria1:="在位次范围内"

    ws.Columns("J:K").Hidden = True

    MsgBox "专业和位次筛选完成！", vbInformation
End Sub

