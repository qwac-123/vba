Attribute VB_Name = "���ϱ�������"
Sub ���ϱ�������()
'
' ���ϱ������� Macro
'x = Application.GetOpenFilename(FileFilter:="Excel�ļ� (*.xls; *.xlsx),*.xls; *.xlsx,�����ļ�(*.*),*.*", _
       Title:="Excelѡ��", MultiSelect:=True) 'ѡ��Ҫ���ϲ��Ĳ�

'
Dim str
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
Set zb = ActiveWorkbook
Set zsh = ActiveSheet
ChDir "\\192.168.0.223\��������\10�����Ϲ���\���α���\����15\2017��"
x = Application.GetOpenFilename(FileFilter:="Excel�ļ� (*.xls; *.xlsx),*.xls; *.xlsx,�����ļ�(*.*),*.*", _
       Title:="Excelѡ��", MultiSelect:=True) 'ѡ��Ҫ���ϲ��Ĳ�
Dim voy '��¼���κ�
If Not IsArray(x) Then '�������ȡ���ͽ���
    GoTo endsub
End If

diyici = True
For Each x1 In x
    If InStr(5, x1, "ȼ����") Then
        GoTo kaishi
    Else
        MsgBox "���ȼ���ϱ���"
        GoTo endsub
    End If
kaishi:
Workbooks.Open (x1)
    Set w = Workbooks.Open(x1)
    Set wsh = w.Sheets("ȼ�ͱ���")
     voy = Mid(w.Name, InStr(11, w.Name, "V") + 1, 4)
    If diyici Then
        wsh.Range("A36:C38,A40:c40").Copy zsh.Cells(1, 2)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "ȼ") - 1)
        diyici = False
    Else
        rowzbwei = zsh.UsedRange.SpecialCells(xlCellTypeLastCell).Row + 1
        If Len(wsh.Range("b38").Text & wsh.Range("c38").Text) = 0 Then
            wsh.Range("A40:C40").Copy zsh.Cells(rowzbwei, 2)
        Else
            wsh.Range("A38:C38,A40:C40").Copy zsh.Cells(rowzbwei, 2)
        End If
        zsh.Cells(rowzbwei, 1) = voy
    End If
    'voy = voy + 1
w.Close
Next x1
Range("b3") = "�ϴ�rob"
Range("b3").Select
ActiveWindow.FreezePanes = True
For i = 4 To Range("b4").End(xlDown).Row
    str = Cells(i, 2).Text
    If InStr(1, str, "�����μ�") Then
        Cells(i, 2) = "+"
    Else
        Cells(i, 2) = "end"
    End If
Next
endsub:

Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
