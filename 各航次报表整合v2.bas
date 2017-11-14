Attribute VB_Name = "各航次报表整合v2"
'
Dim hangciDiYiCi As Boolean
Dim ranrunDiYiCi As Boolean
Dim w As Object
Dim wsh As Object

Dim i As Integer

Dim rng1 As Object
Dim rng2 As Object
Dim rng3 As Object
Dim rgang As Integer '港口所在行
Dim rtou
Dim rwei
Dim rowzbwei
Dim rtouwei
Dim str

Dim rowXiJieTou As Variant
Dim rowXiJieWei

Dim voy '记录航次号

Dim dakaibaobiao As Variant
Dim baobiao As String

Dim zb As Object
Dim zsh As Object

Dim zuoweiuzihoudejieweimeishenmeyisi
Sub 航次报表统一整合()
 'v1.0可以整合航次报表和燃润料报表到一张表里
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0

ChDir "\\192.168.0.223\航运在线\10、油料管理部\航次报表\鼎衡15\2017年"
dakaibaobiao = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿

If Not IsArray(dakaibaobiao) Then '如果点了取消就结束
    Exit Sub
End If

ranrunDiYiCi = True
hangciDiYiCi = True

Set zb = ActiveWorkbook
Set zsh = ActiveSheet

For Each baobiao In dakaibaobiao
    If InStr(5, baobiao, "燃") = 0 Then '这是航次报表
        Call 航次报表整合
    Else '这是燃润料报表
        Call 燃润料报表整合
    End If
Next baobiao


'处理航次区域
    zsh.Columns("c:d").NumberFormatLocal = "ddmmyyhhmm"
    Columns("A:A").ColumnWidth = 4
    Columns("B:B").ColumnWidth = 17.35
    Columns("C:C").ColumnWidth = 9.5
    Columns("D:D").ColumnWidth = 9.5
    Columns("e:i").ColumnWidth = 5.4
    Rows.RowHeight = 12

'处理燃润料区域
Range("L2").Select
ActiveWindow.FreezePanes = True
For i = 2 To Range("L2").End(xlDown).Row
    str = Cells(i, 12).Text
    If InStr(1, str, "本航次加") Then
        Cells(i, 12) = "+"
    Else
        Cells(i, 12) = "end"
    End If
Next
    Columns("k").ColumnWidth = 4
    Columns("l").ColumnWidth = 5
    Columns("m:n").ColumnWidth = 5.88


Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Function 燃润料报表整合()
'v2.0
' 油料报表整合 Macro
'x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿

'
Workbooks.Open (baobiao)
    Set w = Workbooks.Open(baobiao)
    Set wsh = w.Sheets("燃油报表")
     voy = Mid(w.Name, InStr(11, w.Name, "V") + 1, 4)
    If ranrunDiYiCi Then
        wsh.Range("A38:c38").Copy zsh.Cells(1, 2)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "燃") - 1)
        ranrunDiYiCi = False
    End If
    rowzbEnd = zsh.Cells(66666, 2).End(xlUp).Row + 1
    If Len(wsh.Range("b40").Text & wsh.Range("c40").Text) = 0 Then '判断本航次加装这一行是否有加油
        wsh.Range("A42:C42").Copy zsh.Cells(rowzbwei, 2) '只复制航次末结存
    Else
        wsh.Range("A40:C40,A42:C42").Copy zsh.Cells(rowzbwei, 2) '本航次加装和航次末结存
    End If
    zsh.Cells(rowzbwei, 1) = voy
w.Close

Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Function
Function 航次报表整合()
'v1.171114 最后调整了格子大小

    If hangciDiYiCi Then
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
        hangciDiYiCi = False
    Else
        rowzbEnd = zsh.Cells(66666, 5).End(xlUp).Row + 1
        rowXiJieTou = zhaotou() '细节的开头位置
        rowXiJieWei = zhaowei() '细节的最后一条位置
        Set rng1 = wsh.Range(Cells(8, 1), Cells(rgang, 3)) '靠离泊时间区域
        Set rng2 = wsh.Range(Cells(rowXiJieTou, 1), Cells(rowXiJieWei, 3))  '靠离泊细节区域
        Set rng3 = wsh.Range(Cells(rowXiJieTou, 5), Cells(rowXiJieWei, 12))  '细节区域原因
        Union(rng1, rng2).Copy zsh.Cells(rowzbEnd, 2)
        rng3.Copy zsh.Cells(rowzbwei + rgang - 7, 5)
        zsh.Cells(rowzbEnd, 1) = voy
    End If
    'voy = voy + 1
w.Close

End Function
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

