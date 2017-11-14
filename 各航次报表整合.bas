Attribute VB_Name = "各航次报表整合"
'
Dim hangciDiYiCi As Boolean
Dim ranrunDiYiCi As Boolean
Dim openedOil As Boolean
Dim openedVoy As Boolean
Dim w As Object
Dim wsh As Object

Dim i As Integer

Dim rng1 As Object
Dim rng2 As Object
Dim rng3 As Object
Dim rowGangKou 'As Integer '港口所在行
Dim rtou
Dim rwei
Dim rowzbEnd
Dim rtouwei
Dim str As String
Dim shipNum
Dim shipName As String
Dim fileDir
Dim rowXiJieHead As Variant
Dim rowXiJieEnd

Dim voy '记录航次号

Dim dakaibaobiao As Variant
Dim baobiao
Dim zb As Object
Dim zsh As Object
Dim zuoweiuzihoudejieweimeishenmeyisi
Sub yeeee()
Debug.Print "\\192.168.0.223\航运在线\10、油料管理部\航次报表\鼎衡10\" & Year(Date) & "年"
End Sub
Sub 航次报表统一整合()
'v1.4 增加了开头提示清除表
'v1.3 修改了冻结拆分窗格部分
'v1.2 增加了船名输入以便于选择报表文件夹
'v1.1 '增加了判断是否打开过油料表oil 和航次表voy
 'v1.0可以整合航次报表和燃润料报表到一张表里
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
Do
If Len(shipNum) > 2 Then
    MsgBox "请输入正确的数字"
End If
shipNum = InputBox("请输入船名数字，如鼎衡10就输入10", "船名数字", "10")
Loop While Len(shipNum) > 2
If shipNum = "" Then
    Exit Sub
End If
Select Case shipNum
    Case 17
        shipName = "鼎衡17（万年青）"
    Case 18
        shipName = "鼎衡18（常春藤）"
    Case 32
        shipName = "建兴32"
    Case Else
        shipName = "鼎衡" & shipNum
End Select
fileDir = _
"\\192.168.0.223\航运在线\10、油料管理部\航次报表\" _
& shipName & "\" & Year(Date) & "年"
ChDir fileDir
dakaibaobiao = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿

If Not IsArray(dakaibaobiao) Then '如果点了取消就结束
    Exit Sub
End If

ranrunDiYiCi = True
hangciDiYiCi = True
openedOil = False
openedVoy = False
Set zb = ActiveWorkbook
Set zsh = ActiveSheet

If MsgBox("是否清除当前表的内容", vbOKCancel) = vbOK Then
    Cells.Delete '删除当前表的内容
Else
'退出
End If

For Each baobiao In dakaibaobiao
    If InStr(5, baobiao, "燃") = 0 Then '这是航次报表
        Call 航次报表整合
    Else '这是燃润料报表
        Call 燃润料报表整合
    End If
Next baobiao

ActiveWindow.FreezePanes = False
Range("b2").Select
ActiveWindow.FreezePanes = True
'处理航次区域
If openedVoy Then
    zsh.Columns("c:d").NumberFormatLocal = "ddmmyyhhmm"
    Columns("A:A").ColumnWidth = 4
    Columns("B:B").ColumnWidth = 17.35
    Columns("C:C").ColumnWidth = 9.5
    Columns("D:D").ColumnWidth = 9.5
    Columns("e:i").ColumnWidth = 5.4
    Rows.RowHeight = 12
End If
'处理燃润料区域
If openedOil Then
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
End If
Application.ScreenUpdating = 1
Application.DisplayAlerts = 1
End Sub
Function 燃润料报表整合()
'v2.1 增加了判断是否打开过油料表
' 油料报表整合 Macro
'x = Application.GetOpenFilename(FileFilter:="Excel文件 (*.xls; *.xlsx),*.xls; *.xlsx,所有文件(*.*),*.*", _
       Title:="Excel选择", MultiSelect:=True) '选择要被合并的簿

'Workbooks.Open (baobiao)
    Set w = Workbooks.Open(baobiao)
    Set wsh = w.Sheets("燃油报表")
     voy = Mid(w.Name, InStr(11, w.Name, "V") + 1, 4)
    If ranrunDiYiCi Then
        wsh.Range("A38:c38").Copy zsh.Cells(1, 12)
        'zsh.Cells(3, 11) = voy
        zsh.Cells(1, 11) = Mid(w.Name, 1, InStr(3, w.Name, "燃") - 1)
        ranrunDiYiCi = False
    End If
    rowzbEnd = zsh.Cells(66666, 12).End(xlUp).Row + 1
    If Len(wsh.Range("b40").Text & wsh.Range("c40").Text) = 0 Then '判断本航次加装这一行是否有加油
        wsh.Range("A42:C42").Copy zsh.Cells(rowzbEnd, 12) '只复制航次末结存
    Else
        wsh.Range("A40:C40,A42:C42").Copy zsh.Cells(rowzbEnd, 12) '本航次加装和航次末结存
    End If
    zsh.Cells(rowzbEnd, 11) = voy
w.Close
openedOil = True
End Function
Function 航次报表整合()
'v2.1 增加了判断是否打开过航次表
'v1.171114 最后调整了格子大小
    Set w = Workbooks.Open(baobiao)
    Set wsh = w.Sheets("航次报表")
    voy = Mid(w.Name, InStr(6, w.Name, "V") + 1, 4)
    If hangciDiYiCi Then
        rowGangKou = wsh.Cells(8, 3).End(xlDown).Row '靠离泊时间的最后一条位置
        Debug.Print TypeName(rowGangKou)
        rowXiJieHead = zhaotou() '细节的开头位置
        rowXiJieEnd = zhaowei() '细节的最后一条位置
        Set rng1 = wsh.Range(Cells(6, 1), Cells(rowGangKou, 3)) '靠离泊时间区域
        Set rng2 = wsh.Range(Cells(rowXiJieHead, 1), Cells(rowXiJieEnd, 3)) '靠离泊细节区域
        Set rng3 = wsh.Range(Cells(rowXiJieHead, 5), Cells(rowXiJieEnd, 12)) '细节区域原因
        Union(rng1, rng2).Copy zsh.Cells(1, 2)
        rng3.Copy zsh.Cells(rowGangKou - 4, 5)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "航") - 1) 'a1格写船名
        hangciDiYiCi = False
    Else
        rowzbEnd = zsh.Cells(66666, 5).End(xlUp).Row + 1
        rowXiJieHead = zhaotou() '细节的开头位置
        rowXiJieEnd = zhaowei() '细节的最后一条位置
        Set rng1 = wsh.Range(Cells(8, 1), Cells(rowGangKou, 3)) '靠离泊时间区域
        Set rng2 = wsh.Range(Cells(rowXiJieHead, 1), Cells(rowXiJieEnd, 3))  '靠离泊细节区域
        Set rng3 = wsh.Range(Cells(rowXiJieHead, 5), Cells(rowXiJieEnd, 12))  '细节区域原因
        Union(rng1, rng2).Copy zsh.Cells(rowzbEnd, 2)
        rng3.Copy zsh.Cells(rowzbEnd + rowGangKou - 7, 5)
        zsh.Cells(rowzbEnd, 1) = voy
    End If
    'voy = voy + 1
w.Close
openedVoy = True
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
For rowXiJieEnd = rowXiJieHead To 66
    If Cells(rowXiJieEnd, 4) = "" Then
        cishu = cishu + 1
        If cishu > 2 Then
            zhaowei = rowXiJieEnd - cishu
            Exit Function
        End If
    End If
Next
End Function

