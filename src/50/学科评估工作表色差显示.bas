Attribute VB_Name = "学科评估工作表色差显示"
' 学科评估
Sub C_Deal_Classic_By_A()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim visibleRange As Range
    Dim valueDict As Object
    Dim whiteColor As Long
    Dim lightGrayColor As Long

    ' 定义颜色
    whiteColor = RGB(255, 255, 255)  ' 白色
    lightGrayColor = RGB(240, 240, 240) ' 浅灰色
    
    Set ws = ThisWorkbook.Sheets("学科评估") ' 设置工作表名称为“学科评估”
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 清除之前的颜色填充
    ws.Rows("2:" & lastRow).Interior.colorIndex = xlNone

    ' 创建字典来存储 A 列值与颜色的关系
    Set valueDict = CreateObject("Scripting.Dictionary")

    ' 获取筛选后的可见范围
    On Error Resume Next
    Set visibleRange = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' 如果没有可见单元格，则退出
    If visibleRange Is Nothing Then
        MsgBox "没有可见的行可供处理！", vbExclamation
        Exit Sub
    End If

    ' 遍历可见的 A 列（跳过标题行）并记录唯一值
    For Each cell In visibleRange
        If Not IsEmpty(cell.value) Then
            If Not valueDict.Exists(cell.value) Then
                ' 如果是新值，添加到字典
                valueDict.Add cell.value, valueDict.Count Mod 2  ' 用于控制颜色
            End If
        End If
    Next cell

    ' 填充颜色
    Dim currentColor As Long
    Dim lastValue As Variant
    Dim isWhite As Boolean

    ' 从第二行开始
    isWhite = False

    For Each cell In visibleRange
        If Not IsEmpty(cell.value) Then
            ' 如果值不同则切换颜色
            If cell.value <> lastValue Then
                isWhite = Not isWhite
            End If
            
            ' 根据当前颜色填充 A 到 C 列
            If isWhite Then
                ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, 3)).Interior.Color = whiteColor
            Else
                ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, 3)).Interior.Color = lightGrayColor
            End If
            
            lastValue = cell.value
        End If
    Next cell
End Sub

