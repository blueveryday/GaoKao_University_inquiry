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

    ' ����Ĭ�Ϲؼ��֣�ʹ�ö��ŷָ�
    defaultKeywords = "����,�����,���,����,����,����,��е,����,����,˰,��ѧ,����,��,����,ͨ��,�˹�,����,���,��ͨ,����,�Զ���,������,��ȫ,���,���,����,����,������"
    
    ' ����Ĭ����ֵ��Χ
    defaultRange = "2000,5000"
    
    ' ���ù�����
    Set ws = ThisWorkbook.Sheets("רҵ������")
    
    ' ����ؼ��֣�ʹ�ö��ŷָ�
    keywordInput = InputBox("������ؼ��֣��ö��ŷָ�:", "����ؼ���", defaultKeywords)
    
    ' ����û�ȡ����������˳��ӹ���
    If keywordInput = "" Then Exit Sub
    
    ' ���ؼ����ַ������Ϊ����
    keywordArray = Split(keywordInput, ",")
    
    ' ��ȡ��ֵ��Χ����
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
    
    ' �������ɸѡ
    ws.AutoFilterMode = False
    
    ' ����ɸѡ��Χ
    Set rngF = ws.Range("F2", ws.Cells(ws.Rows.Count, "F").End(xlUp)) ' ���� F1 �Ǳ�����
    Set rngH = ws.Range("H2", ws.Cells(ws.Rows.Count, "H").End(xlUp)) ' ���� H1 �Ǳ�����
    
    ' �� H �н�����������
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
    
    ' ������������ݣ����� J �к� K ��Ϊ�����У�
    ws.Range("J:K").ClearContents
    
    ' �ڱ��������ø����еı���
    ws.Range("J1").value = "��ѡרҵ"
    ws.Range("K1").value = "λ������"
    
    ' �������������ڱ�ǰ����ؼ��ֵ���
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
            cell.Offset(0, 4).value = "����" ' �����д�� J ��
        Else
            cell.Offset(0, 4).value = "������"
        End If
    Next cell
    
    ' �������������ڱ����ֵ��Χ����
    For Each cell In rngH
        value = cell.value
        If IsNumeric(value) Then
            If value >= lowerBound And value <= upperBound Then
                cell.Offset(0, 3).value = "��λ�η�Χ��" ' �����д�� K ��
            Else
                cell.Offset(0, 3).value = "����λ�η�Χ��"
            End If
        Else
            cell.Offset(0, 3).value = "����λ�η�Χ��"
        End If
    Next cell
    
    ' Ӧ��ɸѡ�����ų��������������������
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=6, Criteria1:="<>*�������*"
    
    ' Ӧ��ɸѡ��ɸѡ�ؼ���ƥ�����
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=10, Criteria1:="����"
    
    ' Ӧ��ɸѡ��ɸѡ��ֵ��Χ�ڵ���
    ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=11, Criteria1:="��λ�η�Χ��"
    
    ' ���� J �к� K ��
    ws.Columns("J:K").Hidden = True
    
    MsgBox "רҵ��λ��ɸѡ��ɣ�", vbInformation
End Sub


