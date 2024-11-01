Attribute VB_Name = "רҵ�����߹�����ɫ����ʾ"
' רҵ������
Sub B_Deal_Score_By_G()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim visibleRange As Range
    Dim valueDict As Object
    Dim whiteColor As Long
    Dim lightGrayColor As Long

    ' ������ɫ
    whiteColor = RGB(255, 255, 255)  ' ��ɫ
    lightGrayColor = RGB(240, 240, 240) ' ǳ��ɫ
    
    Set ws = ThisWorkbook.Sheets("רҵ������") ' ���ù���������Ϊ��רҵ�����ߡ�
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ' ���֮ǰ����ɫ���
    ws.Rows("2:" & lastRow).Interior.colorIndex = xlNone

    ' �����ֵ����洢 G ��ֵ����ɫ�Ĺ�ϵ
    Set valueDict = CreateObject("Scripting.Dictionary")

    ' ��ȡɸѡ��Ŀɼ���Χ
    On Error Resume Next
    Set visibleRange = ws.Range("G2:G" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' ���û�пɼ���Ԫ�����˳�
    If visibleRange Is Nothing Then
        MsgBox "û�пɼ����пɹ�����", vbExclamation
        Exit Sub
    End If

    ' �����ɼ��� G �У����������У�����¼Ψһֵ
    For Each cell In visibleRange
        If Not IsEmpty(cell.value) Then
            If Not valueDict.Exists(cell.value) Then
                ' �������ֵ����ӵ��ֵ�
                valueDict.Add cell.value, valueDict.Count Mod 2  ' ���ڿ�����ɫ
            End If
        End If
    Next cell

    ' �����ɫ
    Dim currentColor As Long
    Dim lastValue As Variant
    Dim isWhite As Boolean

    ' �ӵڶ��п�ʼ
    isWhite = False

    For Each cell In visibleRange
        If Not IsEmpty(cell.value) Then
            ' ���ֵ��ͬ���л���ɫ
            If cell.value <> lastValue Then
                isWhite = Not isWhite
            End If
            
            ' ���ݵ�ǰ��ɫ��� A �� I ��
            If isWhite Then
                ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, 9)).Interior.Color = whiteColor
            Else
                ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, 9)).Interior.Color = lightGrayColor
            End If
            
            lastValue = cell.value
        End If
    Next cell
End Sub

