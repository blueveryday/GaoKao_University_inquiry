Attribute VB_Name = "重置筛选"
Sub D_Clear_Filter()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.Name = "专业分数线" Or ws.Name = "学科评估" Then
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
            MsgBox "工作表 '" & ws.Name & "' 的筛选已清除！重置筛选条件。", vbInformation
        Else
            MsgBox "工作表 '" & ws.Name & "' 没有应用筛选。", vbExclamation
        End If
        
        ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=3, Criteria1:="2024" ' C列筛选
        ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=4, Criteria1:="物理类" ' D列筛选
        ws.Range("A1:L" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).AutoFilter Field:=5, Criteria1:="本科批" ' 录取批次筛选
        
    Else
        MsgBox "此功能仅适用于“专业分数线”和“学科评估”工作表！", vbExclamation
    End If
End Sub

