Attribute VB_Name = "各航次报表整合"
'

Dim rowXiJieTou As Variant
Dim rowXiJieWei

Sub 航次报表燃润料报表整合()
'v1.0
' 油料报表整合 Macro
'x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿

'
Dim str
Dim zsh
Dim x
Dim x1
Dim w
Dim wsh
Dim rowzbwei
Dim i
Dim zb As Variant

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
        wsh.Range("A38:c38").Copy zsh.Cells(1, 2)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "燃") - 1)
        diyici = False
    End If
    rowzbwei = zsh.Cells(66666, 2).End(xlUp).Row + 1
    If Len(wsh.Range("b40").Text & wsh.Range("c40").Text) = 0 Then '判断本航次加装这一行是否有加油
        wsh.Range("A42:C42").Copy zsh.Cells(rowzbwei, 2) '只复制航次末结存
    Else
        wsh.Range("A40:C40,A42:C42").Copy zsh.Cells(rowzbwei, 2) '本航次加装和航次末结存
    End If
    zsh.Cells(rowzbwei, 1) = voy
w.Close
Next x1

Range("b2").Select
ActiveWindow.FreezePanes = True
For i = 2 To Range("b2").End(xlDown).Row
    str = Cells(i, 2).Text
    If InStr(1, str, "本航次加") Then
        Cells(i, 2) = "+"
    Else
        Cells(i, 2) = "end"
    End If
Next
    Columns("A:A").ColumnWidth = 4
    Columns("B:B").ColumnWidth = 5
    Columns("C:C").ColumnWidth = 5.88
    Columns("D:D").ColumnWidth = 5.88

Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Sub 航次报表整合()
Dim diyici
Dim rgang '港口所在行
Dim str
Dim rng1
Dim rng2
Dim rng3
Dim rtou
Dim zsh
Dim x
Dim x1
Dim w
Dim wsh
Dim rowzbwei
Dim rtouwei
Dim i
Dim zb As Variant
Dim rwei

Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
Set zb = ActiveWorkbook
Set zsh = ActiveSheet
ChDir "\\192.168.0.223\航运在线\10、油料管理部\航次报表\"
x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="选择航次报表", MultiSelect:=True) '选择要被合并的簿
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
    
    'On Error Resume Next
    Set w = Workbooks.Open(x1)
    Set wsh = w.Sheets("航次报表")
    voy = Mid(w.Name, InStr(8, w.Name, "V") + 1, 4)
    If diyici Then
        rgang = wsh.Cells(8, 3).End(xlDown).Row '靠离泊时间的最后一条位置
        rowXiJieTou = zhaotou() '细节的开头位置
        rowXiJieWei = zhaowei() '细节的最后一条位置
        Set rng1 = wsh.Range(Cells(6, 1), Cells(rgang, 3)) '靠离泊时间区域
        Set rng2 = wsh.Range(Cells(rowXiJieTou, 1), Cells(rowXiJieWei, 3)) '靠离泊细节区域
        Set rng3 = wsh.Range(Cells(rowXiJieTou, 5), Cells(rowXiJieWei, 12)) '细节区域原因
        Union(rng1, rng2).Copy zsh.Cells(1, 2)
        rng3.Copy zsh.Cells(rgang - 4, 5)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "航") - 1) 'a1格写船名
        diyici = False
    Else
        rowzbwei = zsh.Cells(66666, 5).End(xlUp).Row + 1
        rowXiJieTou = zhaotou() '细节的开头位置
        rowXiJieWei = zhaowei() '细节的最后一条位置
        Set rng1 = wsh.Range(Cells(8, 1), Cells(rgang, 3)) '靠离泊时间区域
        Set rng2 = wsh.Range(Cells(rowXiJieTou, 1), Cells(rowXiJieWei, 3))  '靠离泊细节区域
        Set rng3 = wsh.Range(Cells(rowXiJieTou, 5), Cells(rowXiJieWei, 12))  '细节区域原因
        Union(rng1, rng2).Copy zsh.Cells(rowzbwei, 2)
        rng3.Copy zsh.Cells(rowzbwei + rgang - 7, 5)
        zsh.Cells(rowzbwei, 1) = voy
    End If
    'voy = voy + 1
w.Close
Next x1
    zsh.Columns("c:d").NumberFormatLocal = "ddmmyyhhmm"
   

Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Function zhaotou()
Dim gezi
For Each gezi In Range("a25:a55") '找到开头的位置
    If gezi = "（纯装卸货时间、补给、抛锚等待、靠泊作业准备时间）" Then
        zhaotou = gezi.Row + 1
        
        'Debug.Print gezi.Row
        Exit Function
    End If
Next gezi
End Function
Function zhaowei()
Dim cishu
Dim i
cishu = 0
For rowXiJieWei = rowXiJieTou To 66
    If Cells(rowXiJieWei, 4) = "" Then
        cishu = cishu + 1
        If cishu > 2 Then
            zhaowei = rowXiJieWei - cishu
            Exit Function
        End If
    End If
Next
End Function

