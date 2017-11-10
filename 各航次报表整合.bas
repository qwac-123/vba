Attribute VB_Name = "各航次报表整合"
Sub 油料报表整合()
'
' 油料报表整合 Macro
'x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿

'
Dim str
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
Set zb = ActiveWorkbook
Set zsh = ActiveSheet
ChDir "\\192.168.0.223\航运在线\10、油料管理部\航次报表\鼎衡15\2017年"
x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿
Dim voy '记录航次号
If Not IsArray(x) Then '如果点了取消就结束
    Exit Sub
End If

diyici = True
For Each x1 In x
    If InStr(5, x1, "燃") = 0 Then
        MsgBox "请打开燃润料报表"
        Exit Sub
    End If

Workbooks.Open (x1)
    Set w = Workbooks.Open(x1)
    Set wsh = w.Sheets("燃油报表")
     voy = Mid(w.Name, InStr(11, w.Name, "V") + 1, 4)
    If diyici Then
        wsh.Range("A36:C38,A40:c40").Copy zsh.Cells(1, 2)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "燃") - 1)
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
Range("b3") = "上次rob"
Range("b3").Select
ActiveWindow.FreezePanes = True
For i = 4 To Range("b4").End(xlDown).Row
    str = Cells(i, 2).Text
    If InStr(1, str, "本航次加") Then
        Cells(i, 2) = "+"
    Else
        Cells(i, 2) = "end"
    End If
Next


Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Sub 航次报表整合()

Dim str
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
Set zb = ActiveWorkbook
Set zsh = ActiveSheet
ChDir "\\192.168.0.223\航运在线\10、油料管理部\航次报表\鼎衡7\2017年"
x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿
Dim voy '记录航次号
If Not IsArray(x) Then '如果点了取消就结束
    Exit Sub
End If

diyici = True
For Each x1 In x
    If InStr(5, x1, "燃") > 0 Then
        MsgBox "请打开航次报表"
        Exit Sub
    End If
    
    On Error Resume Next
    Debug.Print x1
    Set w = Workbooks.Open(x1)
    Set wsh = w.Sheets("航次报表")
     voy = Mid(w.Name, InStr(8, w.Name, "V") + 1, 4)
    If diyici Then
        rTouWei = zhao()
        rgang = wsh.Cells(8, 3).End(xlDown).Row
        rwei = wsh.Cells(41, 3).End(xlDown).Row
        Set rng1 = wsh.Range(Cells(6, 1), Cells(rgang, 3))
        Set rng2 = wsh.Range(Cells(rTouWei, 1), Cells(rwei, 3))
        Union(rng1, rng2).Copy zsh.Cells(1, 2)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "航") - 1) 'a1格写船名
        diyici = False
    Else
        rowzbwei = zsh.UsedRange.SpecialCells(xlCellTypeLastCell).Row + 1
        rgang = wsh.Cells(8, 3).End(xlDown).Row
        rwei = wsh.Cells(41, 3).End(xlDown).Row
        Set rng1 = wsh.Range(Cells(8, 1), Cells(rgang, 3))
        Set rng2 = wsh.Range(Cells(rTouWei, 1), Cells(rwei, 3))
        Union(rng1, rng2).Copy zsh.Cells(rowzbwei, 2)
       
        zsh.Cells(rowzbwei, 1) = voy
    End If
    'voy = voy + 1
w.Close
Next x1
    zsh.Columns("c:d").NumberFormatLocal = "ddmmyyhhmm"
   

Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Function zhao()
For Each gezi In Range("a35:a40")
    If gezi = "（纯装卸货时间、补给、抛锚等待、靠泊作业准备时间）" Then
        zhao = gezi.Row + 2
        Exit Function
    End If
Next

End Function
