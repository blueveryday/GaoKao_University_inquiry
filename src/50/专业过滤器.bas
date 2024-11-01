Attribute VB_Name = "רҵ������"
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
    
    defaultKeywords = "����,�����,���,����,����,����,��е,����,����,˰,��ѧ,����,����,ͨ��,�˹�,����,���,��ͨ,����,�Զ���,������,��ȫ,���,���,����,����,������,����"
    defaultRange = "2900,6000"
    excludeKeywords = "����,����,����"

    Set ws = ThisWorkbook.Sheets("רҵ������")

    yearInput = InputBox("������Ҫɸѡ��������ݣ����磺2023��2024��:", "�����������", "2024")
    If yearInput = "" Then Exit Sub

    ws.AutoFilterMode = False
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:=yearInput

    defaultCategory = "������"
    categoryInput = InputBox("������Ҫɸѡ����������� �� ��ʷ�ࣩ:", "�������", defaultCategory)
    If categoryInput = "" Then Exit Sub
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=4, Criteria1:=categoryInput

    defaultBatch = "������"
    batchInput = InputBox("������Ҫɸѡ�����Σ���������������ǰ����������ǰ��B�Σ�:", "��������", defaultBatch)
    If batchInput = "" Then Exit Sub
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=5, Criteria1:=batchInput

    Set rngF = ws.Range("F2", ws.Cells(ws.Rows.Count, "F").End(xlUp))
    Set rngH = ws.Range("H2", ws.Cells(ws.Rows.Count, "H").End(xlUp))

    keywordInput = InputBox("������ؼ��֣��ö��ŷָ�:", "����ؼ���", defaultKeywords)
    If keywordInput = "" Then Exit Sub
    keywordArray = Split(keywordInput, ",")

    excludeKeywordInput = InputBox("������Ҫ�ų���רҵ���ƹؼ��֣��ö��ŷָ�:", "�����ų��ؼ���", excludeKeywords)
    If excludeKeywordInput = "" Then excludeKeywordArray = Split(excludeKeywords, ",") Else excludeKeywordArray = Split(excludeKeywordInput, ",")

    If excludeKeywordInput = "" Then Exit Sub

    userInput = InputBox("��������ֵ��Χ���ö��ŷָ������磺2000,3000��:", "������ֵ��Χ", defaultRange)
    If userInput = "" Then Exit Sub
    On Error Resume Next
    lowerBound = CDbl(Split(userInput, ",")(0))
    upperBound = CDbl(Split(userInput, ",")(1))
    On Error GoTo 0
    If IsEmpty(lowerBound) Or IsEmpty(upperBound) Then
        MsgBox "��Ч����ֵ��Χ���롣", vbExclamation
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
    ws.Range("J1").value = "��ѡרҵ"
    ws.Range("K1").value = "λ������"

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
            cell.Offset(0, 4).value = "����"
        Else
            cell.Offset(0, 4).value = "������"
        End If
    Next cell

    For Each cell In rngF
        For i = LBound(excludeKeywordArray) To UBound(excludeKeywordArray)
            Dim excludeKeyword As String
            excludeKeyword = Trim(excludeKeywordArray(i))
            If InStr(cell.value, excludeKeyword) > 0 Then
                cell.Offset(0, 4).value = "�ų�"
                Exit For
            End If
        Next i
    Next cell

    For Each cell In rngH
        If IsNumeric(cell.value) Then
            value = cell.value
            If value >= lowerBound And value <= upperBound Then
                cell.Offset(0, 3).value = "��λ�η�Χ��"
            Else
                cell.Offset(0, 3).value = "����λ�η�Χ��"
            End If
        Else
            cell.Offset(0, 3).value = "����λ�η�Χ��"
        End If
    Next cell

    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=6, Criteria1:="<>*�������*"
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=10, Criteria1:="����"
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=11, Criteria1:="��λ�η�Χ��"

    ws.Columns("J:K").Hidden = True

    MsgBox "רҵ��λ��ɸѡ��ɣ�", vbInformation
End Sub

