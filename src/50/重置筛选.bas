Attribute VB_Name = "����ɸѡ"
Sub D_Clear_Filter()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.Name = "רҵ������" Or ws.Name = "ѧ������" Then
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
            MsgBox "������ '" & ws.Name & "' ��ɸѡ�����������ɸѡ������", vbInformation
        Else
            MsgBox "������ '" & ws.Name & "' û��Ӧ��ɸѡ��", vbExclamation
        End If
        
        ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:="2024" ' C��ɸѡ
        ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=4, Criteria1:="������" ' D��ɸѡ
        ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=5, Criteria1:="������" ' ¼ȡ����ɸѡ
        
    Else
        MsgBox "�˹��ܽ������ڡ�רҵ�����ߡ��͡�ѧ��������������", vbExclamation
    End If
End Sub

