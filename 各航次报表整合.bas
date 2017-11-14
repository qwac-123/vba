Attribute VB_Name = "各航次报表整合"
Option Explicit

Dim hangciDiYiCi As Boolean
Dim ranrunDiYiCi As Boolean
Dim openedOil As Boolean
Dim openedVoy As Boolean

Dim i As Integer

Dim rng1 As Object
Dim rng2 As Object
Dim rng3 As Object

Dim w As Object
Dim wsh As Object
Dim zb As Object
Dim zsh As Object

Dim rowGangKou As Long '港口所在行
Dim rowzbEnd As Long
Dim rowXiJieHead As Long '行号数据类型是long
Dim rowXiJieEnd As Long

Dim str As String '单元格内容
Dim shipNum As String '从input里得到的，都是string
Dim shipName As String
Dim fileDir As String '文件夹位置
Dim voy As String '记录航次号

Dim dakaibaobiao As Variant ' 在VBA中，对于For Each m In a，若a是数组，m只能声明为variant 变量，这是语法决定的。
Dim baobiao As Variant '同上

Dim zuoweiuzihoudejieweimeishenmeyisi
Sub 航次报表统一整合()
'v1.5 重做了船名输入选择
'v1.4 增加了开头提示清除表
'v1.3 修改了冻结拆分窗格部分
'v1.2 增加了船名输入以便于选择报表文件夹
'v1.1 '增加了判断是否打开过油料表oil 和航次表voy
 'v1.0可以整合航次报表和燃润料报表到一张表里
Application.ScreenUpdating = 0
Application.DisplayAlerts = 0
Do
Debug.Print shipNum
If shipNum = "     " Then
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
'v2.2 修改了变量，增加找FO:功能
'v2.1 增加了判断是否打开过油料表
'v2.0 从原来的sub改为sub航次报表统一整合()下的一个function
'v1.0 油料报表整合 Macro
Dim rngGezi As Object
Dim rngOilHead As Object
Dim rngOilAdd As Object
Dim rngOilEnd As Object
    Set w = Workbooks.Open(baobiao)
    Set wsh = w.Sheets("燃油报表")
    voy = Mid(w.Name, InStr(11, w.Name, "V") + 1, 4)
For Each rngGezi In Range("b36:b44")
    If rngGezi = "FO:" Then
        rngOilHead = Range(Cells(rngGezi.Row, 2), Cells(rngGezi.Row, 3))
    End If
Next rngGezi
rowOilAdd = Range(Cells(rngGezi.Row + 2, 2), Cells(rngGezi.Row + 2, 3))
rowOilEnd = Range(Cells(rngGezi.Row + 4, 2), Cells(rngGezi.Row + 4, 3))
    If ranrunDiYiCi Then
        rngOilHead.Copy zsh.Cells(1, 12)
        zsh.Cells(1, 11) = Mid(w.Name, 1, InStr(3, w.Name, "燃") - 1)
        ranrunDiYiCi = False
    End If
    rowzbEnd = zsh.Cells(66666, 12).End(xlUp).Row + 1
    If Len(wsh.Range("b40").Text & wsh.Range("c40").Text) = 0 Then '判断本航次加装这一行是否有加油
        rowOilEnd.Copy zsh.Cells(rowzbEnd, 12) '只复制航次末结存
    Else
        Union(rowOilAdd, rowOilEnd).Copy zsh.Cells(rowzbEnd, 12) '本航次加装和航次末结存
    End If
    zsh.Cells(rowzbEnd, 11) = voy
w.Close
openedOil = True
End Function
Function 航次报表整合()
'v2.2 现在只选中可见单元格
'v2.1 增加了判断是否打开过航次表
'v2.0 从原来的sub改为sub航次报表统一整合()下的一个function
'v1.171114 最后调整了格子大小
'v1.0 航次报表整合 Macro
    Set w = Workbooks.Open(baobiao)
    Set wsh = w.Sheets("航次报表")
    voy = Mid(w.Name, InStr(6, w.Name, "V") + 1, 4)
    If hangciDiYiCi Then
        rowGangKou = wsh.Cells(8, 3).End(xlDown).Row '靠离泊时间的最后一条位置
        rowXiJieHead = rowZhaoHead() '细节的开头位置
        rowXiJieEnd = rowFindEnd() '细节的最后一条位置
        Set rng1 = wsh.Range(Cells(6, 1), Cells(rowGangKou, 3)) '靠离泊时间区域
        Set rng2 = wsh.Range(Cells(rowXiJieHead, 1), Cells(rowXiJieEnd, 3)).SpecialCells(xlCellTypeVisible) '靠离泊细节区域
        Set rng3 = wsh.Range(Cells(rowXiJieHead, 5), Cells(rowXiJieEnd, 12)).SpecialCells(xlCellTypeVisible) '细节区域原因
        Union(rng1, rng2).Copy zsh.Cells(1, 2)
        rng3.Copy zsh.Cells(rowGangKou - 4, 5)
        zsh.Cells(3, 1) = voy
        zsh.Range("a1") = Mid(w.Name, 1, InStr(3, w.Name, "航") - 1) 'a1格写船名
        hangciDiYiCi = False
    Else
        rowzbEnd = zsh.Cells(66666, 5).End(xlUp).Row + 1
        rowXiJieHead = rowZhaoHead() '细节的开头位置
        rowXiJieEnd = rowFindEnd() '细节的最后一条位置
        Set rng1 = wsh.Range(Cells(8, 1), Cells(rowGangKou, 3)) '靠离泊时间区域
        Set rng2 = wsh.Range(Cells(rowXiJieHead, 1), Cells(rowXiJieEnd, 3)).SpecialCells(xlCellTypeVisible)  '靠离泊细节区域
        Set rng3 = wsh.Range(Cells(rowXiJieHead, 5), Cells(rowXiJieEnd, 12)).SpecialCells(xlCellTypeVisible)  '细节区域原因
        Union(rng1, rng2).Copy zsh.Cells(rowzbEnd, 2)
        rng3.Copy zsh.Cells(rowzbEnd + rowGangKou - 7, 5)
        zsh.Cells(rowzbEnd, 1) = voy
    End If
w.Close
openedVoy = True
End Function
Function rowZhaoHead()
Dim strgezi As String
Dim rngGezi As Object
For Each rngGezi In Range("a25:a55") '找到开头的位置
    If rngGezi.Text = "（纯装卸货时间、补给、抛锚等待、靠泊作业准备时间）" Then '如果是"船位 Location"会导致选中前面30行
        rowZhaoHead = rngGezi.Row + 1
        Exit Function
    End If
Next rngGezi
End Function
Function rowFindEnd()
'v1.2 现在可以正确统计连续空行而不是累计空行，并排除隐藏单元格（dh9的）
'
Dim cishu
Dim i
Dim rngGezi As Object
cishu = 0
'Range(Cells(rowXiJieHead, 3), Cells(80, 3)).SpecialCells(xlCellTypeVisible).Select '选中可见单元格
For Each rngGezi In Range(Cells(rowXiJieHead, 3), Cells(80, 3)).SpecialCells(xlCellTypeVisible)
    rowXiJieEnd = rngGezi.Row
    If Cells(rowXiJieEnd, 4) = "" Then
        cishu = cishu + 1
    Else
        cishu = 0
    End If
    If cishu > 2 Then '如果连续3次
        rowFindEnd = rowXiJieEnd - cishu
        Exit Function
    End If
Next rngGezi
End Function
